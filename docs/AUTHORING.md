# Authoring guide

This guide covers how to add or change lessons, workbooks, assignments, quizzes and exams. All content is plain JavaScript in `assets/js/content/` and loads in the order listed in `index.html`:

```
course.js      course facts, units, grading table, shared data sets and data files
unit1.js       weeks 1–5 (lessons, practice workbooks, assignments)
unit2.js       weeks 6–10
unit3.js       weeks 11–15
quizzes.js     the 12 Monday quizzes
exams.js       the 3 unit exams
```

`quiz-template.js` is a commented, working example of every question type. It is disabled, and `index.html` does not load it.

After any change, run `npm test`. The content suite rebuilds every answer key and fails if one does not score exactly 100%, if a blank submission scores 10% or more, if a formula does not parse, or if the calendar breaks the syllabus rules.

## The pieces

### A lesson (one week)

```js
C.lessons.push({
  week: 7, unit: 2, id: 'week-7',
  title: 'Lookups and keys',
  summary: 'One sentence shown on lesson cards.',
  objectives: ['…', '…', '…'],                  // at least 3
  sections: [{ heading: 'Primary keys', html: '<p>…</p>' }],   // the notes; at least 3
  keyTerms: [['Primary key', 'A column whose value is unique for every row.']],   // at least 4
  practice: [ /* practice workbooks, see below */ ],
  assignment: { /* due Sunday of this week, see below */ },
});
```

Section `html` is trusted, author-written HTML. Use `<code>` for formulas, `<table class="notes-table">` for small tables, and `<div class="callout">` (or `class="callout tip"`) for asides.

### An assessment (practice, assignment, quiz or exam)

All four share one shape:

| Field | Meaning |
|---|---|
| `id` | Unique id. It appears in URLs, receipts and the gradebook. Practice ids only need to be unique within their week. |
| `title`, `intro` | Shown to students. Assignments use `html` for longer instructions. |
| `workbook` | The starting workbook (see below). Omit it for page-only quizzes. |
| `tools` | Which spreadsheet tools appear: a preset name (`basic`, `data`, `organize`, `analyze`, `all`) or a list (see **Tools**). |
| `model(wb, { W, F })` | Optional. Steps the answer key performs before solutions are applied, for work no single solution expresses, such as importing a file or removing duplicates. |
| `questions` | The graded items (see **Question types**). |
| `points` | Assignments: the stated total. Auto-graded points come from the questions. |
| `manualPoints`, `rubric` | Points the teacher scores by hand, and the rubric rows `[criterion, points or 'Teacher check']`. |
| `upload` | Assignments only: `{ accept: '.pdf,.png', label: '…' }` adds a file upload that is packed into the receipt. |
| `category: 'project'` | Counts the assignment in the project category (10%) instead of assignments. |
| `week`, `covers`, `minutes`, `kind` | Quizzes and exams: the week given, the weeks assessed, the time limit, and `'quiz'` or `'exam'`. Exams also set `unit`. |

### A starting workbook

```js
workbook: {
  sheets: [{
    name: 'Sales',
    rows: [['Region', 'Units'], ['East', 10]],    // fills from A1; formulas start with =; row 1 is bolded
    headerRow: false,                              // optional: do not treat row 1 as a header
    cells: { D1: { input: 'Notes', locked: true } },
    formats: { 'B2:B20': 'currency' },             // general number integer currency percent percent1 date longdate text
    styles: { 'A1:B1': { bold: true } },
    widths: { A: 120 },                            // pixels
    lock: 'A1:B1',                                 // true locks every given cell; a range or list locks only those
  }],
  tables: [{ name: 'SalesTbl', sheet: 'Sales', ref: 'A1:B20' }],
  names: { Rate: 'Sales!$D$1' },
  sources: { 'sales_jan.csv': '…csv text…' },      // files Get Data and From Text/CSV can import
  queries: [ /* Power Query definitions */ ],
}
```

Locked cells cannot be edited or sorted. To let students sort the data, lock only the header row.

`course.js` has shared data sets that keep the course consistent. Examples are `D.storeSalesSheet(name, opts)` for the 60-order school store data, `D.survey`, `D.students`, `D.courses`, `D.grades` and `D.catalog`, and the data files in `D.files`. Reusing them means numbers in the notes match the workbooks.

## Question types

Every question has `id`, `type`, `prompt` and `points` (default 1). Workbook questions also take `sheet`, which defaults to the first sheet.

### Answered on the page

| Type | Fields | Grading |
|---|---|---|
| `choice` | `options`, `answer` (index), `multi: true` with `answer: [i, j]`, `explain` | Single answer: right or wrong. Multi: (right picks − wrong picks) ÷ correct options. |
| `number` | `answer`, `tolerance` | Accepts `1,234.5`, `$12` and `15%`. |
| `text` | `accept: ['primary key', /^PK$/i]` | Ignores case, extra spaces and punctuation. Regexes test the raw answer. |
| `match` | `pairs: [[term, definition], …]` | Partial credit per pair. The definitions are shown in a stable shuffled order. |

### Answered in the workbook

| Type | Fields | Checks |
|---|---|---|
| `cell` | `cell`, `solution` or `expect`, `requireFormula`, `functions`, `mustReference`, `forbidFunctions`, `tolerance`, `caseSensitive`, `hint` | The value matches the key. Method rules then take partial credit: a typed value where a formula is required (−50%), a missing function or reference (−25%). |
| `range` | `range`, `solution` (one formula filled across, or an array), plus the `cell` options | Credit per cell. |
| `cellType` | `cell` or `range`, `valueType: 'number' \| 'text' \| 'date'`, `expect`, `checkValue: false` | How each value is stored, for example leading zeros kept as text or real dates. |
| `format` | `range`, `format` (or a list of acceptable formats) | The number format of every cell. |
| `filled` | `range`, `min` | At least `min` cells are filled. Use it for writing the teacher scores by hand. |
| `table` | `range`, `name` | An Excel Table covers the range: half credit if it only overlaps, three quarters if the name differs. |
| `validation` | `range`, `solution: { type: 'list', values }` or `{ type: 'whole' \| 'decimal', min, max }` | A rule covers the range with the right type and settings. |
| `condFormat` | `range`, `solution: { op, value, value2, style }` | `op` is `> < >= <= = <> between contains duplicate top aboveAverage blank`. |
| `sorted` | `range` (with header row), `column`, `order: 'asc' \| 'desc'` | The rows are unchanged as records (nothing split apart) and ordered by the column. |
| `chart` | `chartType` (or a list), `solution: { type, range, title }`, `requireTitle`, `axisFromZero`, `anyRange` | Chart type, data range, title and value-axis minimum. |
| `pivot` | `solution: { sourceRange, rows, values, agg, columns, rowsGroup }` | Row and value fields, summary function (`sum count average min max`), column field, date grouping (`month quarter year`). |
| `query` | `solution: { name, source, steps }`, `steps` (types required; a nested list means any one of them), `requireSource`, `loaded` | Step types present, output row count equal to the key's, the right source, and loaded to a sheet that is up to date. |

### How the answer key is built

1. Start from `workbook`.
2. Run `model(wb)` if there is one.
3. Apply each workbook question's `solution` in this order: query, table, pivot, sorted, format, validation, condFormat, chart, cellType, cell, range.

When `solution` is omitted, the key uses `expect`, then `format` for format questions, `true` for sorted and `name` for table. `solution: null` means the model already produced the result, so the question checks what the model left.

A formula solution such as `'=SUMIF(Sales!A:A,"East",Sales!E:E)'` is evaluated on the data, so changing the data never leaves a stale key. Use `expect` when the answer is a fixed fact.

## Tools

The toolbar shows only what an assessment lists in `tools`:

`format bold fill formulas sort filter textToColumns removeDuplicates import csv table validation condFormat names chart pivot query sheets download reset`

| Preset | Tools |
|---|---|
| `basic` | format, bold, fill, formulas, download, reset |
| `data` | basic + sort, textToColumns, removeDuplicates, import, csv |
| `organize` | basic + sort, filter, table, validation, condFormat, names |
| `analyze` | organize + chart, pivot (no validation) |
| `all` | everything, including Power Query and adding sheets |

## Quizzes made in the browser

The **Quiz builder** (Quiz Templates → Open the quiz builder) makes choice, number, short-answer, matching and workbook-cell questions on a workbook the teacher types or loads from `.xlsx`. It refuses to save a quiz whose key does not score 100%. Quizzes are stored in the browser and can be exported as `.quiz.json` to share or import elsewhere. The exported object has the same shape as a course quiz, so it can be pasted into `quizzes.js`.

## Checklist for a new quiz

1. Copy `quiz-template.js`, set a new `id` and `week`, and write the questions.
2. Set `ENABLED = true` and add a `<script>` tag for the file after `exams.js` in `index.html`.
3. Run `npm run test:content`.
4. Open the quiz in the browser, take it once in the web workbook, and once by uploading the blank template completed in Excel.
