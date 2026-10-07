'use strict';
// Browser tests: the real site in headless Chromium, driven the way a student
// uses it. Skipped when Playwright is not installed (npm install playwright).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { ROOT, X, G, findAssessment } = require('./helpers');

let chromium = null;
try { ({ chromium } = require('playwright')); } catch (e) { /* skipped below */ }
const skip = !chromium && 'Playwright is not installed';

const { serve } = require('../scripts/serve.js');

let server, browser, base;
test.before(async () => {
  if (skip) return;
  server = await serve();
  base = 'http://127.0.0.1:' + server.address().port + '/index.html';
  browser = await chromium.launch();
});
test.after(async () => { if (browser) await browser.close(); if (server) server.close(); });

async function newPage(viewport = { width: 1360, height: 1000 }) {
  const page = await browser.newPage({ viewport, acceptDownloads: true });
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') page.errors.push(m.text()); });
  return page;
}

// The cells a perfect answer types, computed from the assessment's own key.
async function keyEntries(page, id) {
  return page.evaluate((id) => {
    const { Grader: G, Formula: F, Workbook: W, App } = window.DIM;
    const a = App.registry().get(id).a;
    const key = G.buildKey(a);
    const out = [];
    for (const q of a.questions) {
      if (q.type !== 'cell' && q.type !== 'range') continue;
      const sheet = q.sheet || key.sheets[0].name;
      const r = W.normRange(q.cell || q.range);
      for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
        const input = key.getInput(sheet, { row, col });
        out.push({ addr: F.quoteSheet(sheet) + '!' + F.toAddress(row, col), input: String(input) });
      }
    }
    return out;
  }, id);
}

async function typeCells(page, scope, entries) {
  const nb = scope.locator('.wb-namebox');
  for (const { addr, input } of entries) {
    await nb.fill(addr);
    await nb.press('Enter');
    await page.keyboard.type(input);
    await page.keyboard.press('Enter');
  }
}

// Answer every page question correctly through the form controls.
async function answerQuestions(page, scope, a) {
  for (const q of a.questions) {
    const li = scope.locator('li.question[data-q="' + q.id + '"]');
    if (q.type === 'choice') for (const i of [].concat(q.answer)) await li.locator('input').nth(Number(i)).check();
    if (q.type === 'number') await li.locator('input').fill(String(q.answer));
    if (q.type === 'text') await li.locator('input').fill(q.accept.find((x) => typeof x === 'string'));
    if (q.type === 'match') {
      const selects = li.locator('select');
      for (let i = 0; i < q.pairs.length; i++) await selects.nth(i).selectOption(String(i));
    }
  }
}

async function signIn(page) {
  await page.locator('.modal input').first().fill('Test Student');
  await page.getByRole('button', { name: 'Save' }).click();
}

test('every page renders without script errors', { skip }, async () => {
  const page = await newPage();
  const routes = ['', 'lessons', 'work', 'templates', 'builder', 'grades', 'settings', 'help', 'nowhere'];
  for (let w = 1; w <= 15; w++) routes.push('week/' + w);
  for (const r of routes) {
    await page.goto(base + '#/' + r);
    await page.waitForSelector('main, .container');
    assert.ok((await page.locator('#app').innerText()).length > 50, r + ' has content');
  }
  for (const id of ['hw-1', 'hw-7', 'hw-14', 'quiz-2', 'exam-1']) {
    await page.goto(base + '#/' + (id.startsWith('hw') ? 'assignment' : id.startsWith('exam') ? 'exam' : 'quiz') + '/' + id);
    await page.waitForTimeout(100);
  }
  assert.deepEqual(page.errors, []);
  await page.close();
});

test('home page: green banner with weeks grouped by unit, and simple navigation', { skip }, async () => {
  const page = await newPage();
  await page.goto(base);
  const nav = await page.locator('header nav a').allInnerTexts();
  for (const label of ['Home', 'Lessons', 'Assignments & Quizzes', 'Quiz Templates', 'Grades']) assert.ok(nav.some((t) => t.includes(label)), label);
  const weeks = await page.locator('.banner a[href^="#/week/"]').count();
  assert.equal(weeks, 15);
  const bg = await page.locator('.banner').evaluate((el) => getComputedStyle(el).backgroundColor + ' ' + getComputedStyle(el).backgroundImage);
  assert.match(bg, /rgb\(\s*\d+,\s*1[0-9]{2},\s*\d+\)|gradient/);
  await page.close();
});

test('practice workbook: typing the right formulas and checking scores 100%', { skip }, async () => {
  const page = await newPage();
  await page.goto(base + '#/week/3/practice-p3a');
  const block = page.locator('#practice-p3a');
  await block.getByRole('button', { name: 'Open workbook' }).click();
  await typeCells(page, block, await keyEntries(page, 'week-3/p3a'));
  await block.getByRole('button', { name: 'Check my work' }).click();
  await assert.doesNotReject(block.locator('.score-banner').waitFor());
  assert.match(await block.locator('.score-banner').innerText(), /100%/);
  assert.deepEqual(page.errors, []);
  await page.close();
});

test('quiz: sign in, answer everything, submit, and it is graded 100% and recorded', { skip }, async () => {
  const page = await newPage();
  await page.goto(base + '#/quiz/quiz-4');
  await page.getByTestId('start').click();
  await signIn(page);
  if (await page.getByTestId('start').isVisible().catch(() => false)) await page.getByTestId('start').click();
  const ws = page.locator('.workspace');
  await ws.waitFor();
  await answerQuestions(page, ws, findAssessment('quiz-4'));
  await typeCells(page, ws, await keyEntries(page, 'quiz-4'));
  await page.getByTestId('submit').click();
  await page.getByRole('button', { name: 'Submit', exact: true }).last().click();
  await page.locator('.score-banner').first().waitFor();
  assert.match(await page.locator('.score-banner').first().innerText(), /100%/);
  const stored = await page.evaluate(() => window.DIM.App.state().submissions.map((s) => [s.assessmentId, s.percent, s.mode]));
  assert.ok(stored.some(([id, p]) => id === 'quiz-4' && p === 100), JSON.stringify(stored));
  await page.goto(base + '#/grades');
  assert.match(await page.locator('#app').innerText(), /quiz|Quiz/);
  assert.deepEqual(page.errors, []);
  await page.close();
});

test('assignment: uploading a finished .xlsx from Excel is graded automatically', { skip }, async () => {
  const a = findAssessment('hw-3');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'dim-e2e-'));
  const file = path.join(dir, 'hw-3.xlsx');
  fs.writeFileSync(file, X.write(G.buildKey(a)));
  const page = await newPage();
  await page.goto(base + '#/assignment/hw-3');
  await page.getByRole('tab', { name: /Upload from Excel/ }).click();
  await page.locator('.upload-host input[type=file]').setInputFiles(file);
  await page.locator('.upload-host .status-ok').waitFor();
  const ws = page.locator('.workspace');
  await answerQuestions(page, ws, a);
  await page.getByTestId('submit').click();
  await page.locator('.modal').first().waitFor();
  if (await page.locator('.modal input').first().isVisible().catch(() => false)) await signIn(page);
  await page.getByRole('button', { name: 'Submit', exact: true }).last().click();
  await page.getByText('Submitted: 100% auto-graded').waitFor();
  // feedback is shown next to every task in the page that was submitted
  assert.equal(await page.locator('li.task.correct').count(), a.questions.filter((q) => G.isWorkbookQuestion(q)).length);
  assert.ok(await page.locator('.upload-host .status-ok').isVisible(), 'the uploaded file is still listed');
  const sub = await page.evaluate(() => window.DIM.App.state().submissions.at(-1));
  assert.equal(sub.assessmentId, 'hw-3');
  assert.equal(sub.mode, 'upload');
  assert.equal(sub.uploadName, 'hw-3.xlsx');
  assert.equal(sub.autoPercent, 100, JSON.stringify(sub.items.filter((i) => i.status !== 'correct')));
  assert.deepEqual(page.errors, []);
  fs.rmSync(dir, { recursive: true, force: true });
  await page.close();
});

test('quiz builder: a teacher makes a quiz with a workbook cell task, and a student takes it for 100%', { skip }, async () => {
  const page = await newPage();
  await page.goto(base + '#/builder');
  await page.getByLabel('Title').fill('Pop quiz: totals');
  await page.getByRole('button', { name: 'Add a blank workbook' }).click();
  const grid = page.locator('.builder-wb, .wb-host, .card').filter({ has: page.locator('.wb-namebox') }).first();
  await typeCells(page, grid, [{ addr: 'A1', input: 'Qty' }, { addr: 'A2', input: '4' }, { addr: 'A3', input: '6' }]);
  await page.getByRole('button', { name: '+ Multiple choice' }).click();
  let dlg = page.locator('.modal');
  await dlg.getByLabel('Question / task').fill('A primary key must be…');
  await dlg.getByLabel('Options').fill('Unique\nBlank\nA formula');
  await dlg.getByLabel('Correct option number(s)').fill('1');
  await dlg.getByRole('button', { name: 'Add' }).click();
  await page.getByRole('button', { name: '+ Workbook cell' }).click();
  dlg = page.locator('.modal');
  await dlg.getByLabel('Question / task').fill('Total quantity in A4');
  await dlg.getByLabel('Cell').fill('A4');
  await dlg.getByLabel('Answer key').fill('=SUM(A2:A3)');
  await dlg.getByRole('button', { name: 'Add' }).click();
  await page.getByRole('button', { name: 'Save quiz' }).click();
  await page.waitForURL(/#\/templates/);
  const quiz = await page.evaluate(() => window.DIM.App.state().custom.at(-1));
  assert.equal(quiz.title, 'Pop quiz: totals');
  assert.equal(quiz.questions.length, 2);
  // take it
  await page.goto(base + '#/quiz/' + quiz.id);
  await page.getByTestId('start').click();
  await signIn(page);
  const ws = page.locator('.workspace');
  await ws.waitFor();
  await answerQuestions(page, ws, quiz);
  await typeCells(page, ws, [{ addr: 'A4', input: '=SUM(A2:A3)' }]);
  await page.getByTestId('submit').click();
  await page.getByRole('button', { name: 'Submit', exact: true }).last().click();
  await page.locator('.score-banner').first().waitFor();
  assert.match(await page.locator('.score-banner').first().innerText(), /100%/);
  assert.deepEqual(page.errors, []);
  await page.close();
});

test('phone width: no sideways scrolling on the main pages', { skip }, async () => {
  const page = await newPage({ width: 375, height: 800 });
  const routes = ['', 'lessons', 'work', 'templates', 'builder', 'grades', 'settings', 'help', 'quiz/quiz-2', 'exam/exam-1', 'assignment/hw-14'];
  for (let w = 1; w <= 15; w++) routes.push('week/' + w);
  for (const r of routes) {
    await page.goto(base + '#/' + r);
    await page.waitForTimeout(100);
    const over = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    assert.ok(over <= 1, '#/' + r + ' overflows by ' + over + 'px');
  }
  await page.close();
});
