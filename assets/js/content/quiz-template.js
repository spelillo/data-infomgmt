/*
 * quiz-template.js — a starting point for adding your own quiz to the course.
 *
 * HOW TO USE
 *   1. Copy this file (for example to assets/js/content/my-quiz.js).
 *   2. Change the id, week, title and questions below. Delete the question
 *      types you do not need; every type the course supports is shown once.
 *   3. Set ENABLED to true.
 *   4. Load it in index.html after exams.js:
 *        <script src="assets/js/content/my-quiz.js"></script>
 *   5. Run `npm test`. The content tests build the answer key from your
 *      `solution` values and fail if the key does not score exactly 100%,
 *      or if a blank submission earns 10% or more.
 *
 * The full reference for every option is docs/AUTHORING.md. Teachers who do
 * not want to edit files can use the Quiz builder page instead.
 *
 * HOW GRADING WORKS (short version)
 *   The answer key is the starting workbook with every question's `solution`
 *   applied, so a formula solution like '=SUM(B2:B9)' stays correct if you
 *   change the data. Students' work is compared to the key cell by cell, and
 *   structures (tables, charts, PivotTables, queries) are checked by what they
 *   are and what they produce. The same rules grade the in-browser workbook and
 *   an .xlsx uploaded from Excel.
 */
(function (root) {
  'use strict';
  const C = typeof module === 'object' && module.exports ? require('./course.js') : root.DIM.Content;

  // Leave false in the template so it never appears in the course by accident.
  const ENABLED = false;

  const quiz = {
    id: 'quiz-template',         // unique across the course; used in URLs and receipts
    week: 4,                     // the lesson week the quiz belongs to (1–15)
    covers: [3, 4],              // weeks whose material it assesses (shown to students)
    title: 'Template quiz: every question type',
    intro: 'Short instructions shown on the start screen.',
    kind: 'quiz',                // 'quiz' or 'exam' (exams are weighted separately)
    minutes: 15,                 // time limit; the quiz auto-submits when it runs out
    tools: ['format', 'table', 'validation', 'condFormat', 'sort', 'chart', 'pivot', 'query', 'download', 'reset'],

    // The starting workbook. `rows` fills cells from A1; `cells` sets single
    // cells (with options such as locked); `formats` applies number formats.
    workbook: {
      sheets: [
        {
          name: 'Sales',
          rows: [
            ['Region', 'Product', 'Units', 'Price', 'Revenue'],
            ['East', 'Pens', 10, 2.5],
            ['West', 'Paper', 4, 3],
            ['East', 'Paper', 7, 3],
            ['North', 'Pens', 12, 2.5],
          ],
          formats: { 'D2:E5': 'currency' },
          lock: 'A1:E1',           // locked cells cannot be edited: true locks every given cell; here only the header, so rows can be sorted
          widths: { A: 90, B: 90 },
        },
        { name: 'Answers', rows: [['Total revenue', null], ['East revenue', null], ['Order ID', null]] },
      ],
      // Text files a query can import (Get Data). Keys are file names.
      sources: { 'sales.csv': 'Region,Units\nEast,10\nWest,4\nEast,7\n' },
    },

    questions: [
      // ----- questions answered on the page ------------------------------------
      { id: 'q1', type: 'choice', points: 1, prompt: 'Which column should be the primary key of an Orders table?',
        options: ['Customer name', 'Order ID', 'Order date', 'Total'], answer: 1,
        explain: 'A primary key is unique and never blank.' },
      { id: 'q2', type: 'choice', multi: true, points: 2, prompt: 'Select every number format that keeps the stored value unchanged.',
        options: ['Currency', 'Percent', 'Typing the value as text', 'Date'], answer: [0, 1, 3] },
      { id: 'q3', type: 'number', points: 1, prompt: 'How many East orders are there?', answer: 2, tolerance: 0 },
      { id: 'q4', type: 'text', points: 1, prompt: 'What is a column in another table that points at a primary key called?',
        accept: ['foreign key', 'FK', /^foreign\s*keys?$/i] },
      { id: 'q5', type: 'match', points: 2, prompt: 'Match each term to its meaning.',
        pairs: [['Record', 'One row'], ['Field', 'One column'], ['Primary key', 'Uniquely identifies a row']] },

      // ----- questions answered in the workbook -----------------------------------
      // cell: one cell, compared to the key. requireFormula halves credit for a
      // typed value; functions lists required functions ([['A','B']] = A or B).
      { id: 'w1', type: 'cell', sheet: 'Answers', cell: 'B1', points: 2, prompt: 'Total revenue in B1.',
        solution: '=SUM(Sales!E2:E5)', requireFormula: true, functions: ['SUM'], tolerance: 0.005 },
      { id: 'w2', type: 'cell', sheet: 'Answers', cell: 'B2', points: 2, prompt: 'East revenue in B2 with SUMIF or SUMIFS.',
        solution: '=SUMIF(Sales!A2:A5,"East",Sales!E2:E5)', functions: [['SUMIF', 'SUMIFS']],
        hint: 'SUMIF(range, criteria, sum_range)' },
      // range: every cell in the range. A single formula solution is filled down.
      { id: 'w3', type: 'range', sheet: 'Sales', range: 'E2:E5', points: 3, prompt: 'Revenue = Units × Price in E2:E5.',
        solution: '=C2*D2', requireFormula: true },
      // cellType: checks how a value is stored (number, text or date).
      { id: 'w4', type: 'cellType', sheet: 'Answers', cell: 'B3', valueType: 'text', expect: '00042', points: 1,
        prompt: 'Type the order ID 00042 in B3 so the leading zeros are kept.' },
      // format: the number format of a range.
      { id: 'w5', type: 'format', sheet: 'Answers', range: 'B1:B2', format: 'currency', points: 1, prompt: 'Format B1:B2 as currency.' },
      // table: an Excel Table covering the range, optionally with a name.
      { id: 'w6', type: 'table', sheet: 'Sales', range: 'A1:E5', name: 'SalesTbl', points: 1, prompt: 'Format A1:E5 as a table named SalesTbl.' },
      // validation: list, whole or decimal rules.
      { id: 'w7', type: 'validation', sheet: 'Sales', range: 'A2:A5', points: 1, prompt: 'Only allow East, West, North or South in A2:A5.',
        solution: { type: 'list', values: ['East', 'West', 'North', 'South'] } },
      // condFormat: op is one of > < >= <= = <> between contains duplicate top aboveAverage blank.
      { id: 'w8', type: 'condFormat', sheet: 'Sales', range: 'C2:C5', points: 1, prompt: 'Highlight units over 8 in green.',
        solution: { op: '>', value: 8, style: 'green' } },
      // sorted: rows kept together and ordered by one column.
      { id: 'w9', type: 'sorted', sheet: 'Sales', range: 'A1:E5', column: 'C', order: 'desc', points: 1, prompt: 'Sort the data by Units, largest first.' },
      // chart: type (column, bar, line, pie, scatter), optional title and axis checks.
      { id: 'w10', type: 'chart', sheet: 'Sales', chartType: ['column', 'bar'], points: 2, prompt: 'A titled column chart of units by region, axis from 0.',
        solution: { type: 'column', range: 'A1:A5', title: 'Units by region' }, anyRange: true, requireTitle: true, axisFromZero: true },
      // pivot: rows, values, summary function; rowsGroup groups dates by month/quarter/year.
      { id: 'w11', type: 'pivot', sheet: 'Sales', points: 2, prompt: 'A PivotTable of total units by region.',
        solution: { sourceRange: 'A1:E5', rows: 'Region', values: 'Units', agg: 'sum' } },
      // query: Power Query steps and output. steps lists the step types that must
      // appear; loaded requires Close & Load. In Excel the M code is read and mapped.
      { id: 'w12', type: 'query', points: 2, prompt: 'A query on sales.csv that removes duplicates, loaded to a sheet.',
        solution: { name: 'Clean', source: { kind: 'file', file: 'sales.csv' }, steps: [{ type: 'removeDuplicates', columns: [] }] },
        steps: ['removeDuplicates'], loaded: true },
      // filled: something written in a range (for reflections scored by hand).
      { id: 'w13', type: 'filled', sheet: 'Answers', range: 'A6:A8', min: 1, points: 1, prompt: 'Write one sentence about the data in A6.' },
    ],
  };

  if (ENABLED) C.quizzes.push(quiz);
  if (typeof module === 'object' && module.exports) module.exports = quiz;
})(typeof globalThis !== 'undefined' ? globalThis : this);
