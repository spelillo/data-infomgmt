# Intro to Data and Information Management

The course site for a one-semester, 15-week, 12th-grade course on how organizations collect, organize and use data. Students learn hands-on in Excel, then make their work repeatable with Power Query and Power BI.

Open `index.html` and the whole course runs in the browser. There is no build step, no server and no account to set up.

- **Lessons:** detailed notes for every week, with objectives, key terms and worked examples.
- **Interactive Excel workbooks:** a spreadsheet in the page with real formulas, tables, validation, conditional formatting, charts, PivotTables and a Power Query editor.
- **Assignments, quizzes and exams:** auto-graded the moment they are submitted, whether the student works in the browser or uploads an `.xlsx` saved in Excel.
- **Quiz templates:** blank `.xlsx` workbooks for every graded item, plus a quiz builder for writing new auto-graded quizzes.
- **Grades:** each student's weighted course grade, and a teacher gradebook that re-grades every handed-in receipt from the work inside it.

## Using the site

| Page | What it is for |
|---|---|
| **Home** | Course overview, this week's work, and the grading breakdown. The green banner under the header links every week, grouped by unit. |
| **Lessons** | All 15 weeks. Each week has notes, key terms, practice workbooks with **Check my work**, its assignment, and its quiz or exam. |
| **Assignments & Quizzes** | Everything that is graded, with due dates. Assignments are due Sunday at 11:59 pm, quizzes are on Monday, and exams are on the Friday of weeks 5, 10 and 15. |
| **Quiz Templates** | Blank Excel templates for every graded item, the quiz builder, and the authoring template. |
| **Grades** | A student's submissions and course grade. Teachers import receipts here to build the gradebook. |

**For students.** Sign in with your name, which is stored only in your browser. Work in the browser workbook, or download the blank template, finish it in Excel, and upload the `.xlsx`. After you submit, download the receipt and turn it in where your teacher asks.

**For teachers.** Set the first day of class in **Settings** so due dates line up. Collect receipt files and import them on **Grades**. Each one is re-graded from the workbook inside it, and a mismatch with the score the student's browser reported is flagged. Score the rubric items, then export the gradebook as CSV. Settings also chooses whether a student's first or best attempt counts.

Progress is saved in the browser's `localStorage`. **Settings** can back it up to a file, restore it, or erase it.

## Calendar and grading

| Category | Weight |
|---|---|
| Weekly assignments (weeks 1–15) | 30% |
| Monday quizzes (12, lowest dropped) | 15% |
| Exam 1, Exam 2, Exam 3 | 15% each |
| Final dashboard project (week 14) | 10% |

There is no quiz in week 1, and none on the Monday after an exam (weeks 6 and 11). Each exam covers only its own unit. Primary and foreign keys are taught in week 7, alongside lookups.

## How grading works

Every assessment's answer key is built from the starting workbook plus each question's solution. A student's work is then compared to that key cell by cell and structure by structure, so the key is always consistent with the data. The test suite proves that every key scores exactly 100% and that a blank submission scores under 10%.

Uploads from desktop Excel are graded by the same rules. The reader understands formulas (including shared and structured references), number formats, Tables, data validation, conditional formatting, charts, PivotTables (with date grouping), and Power Query queries. Queries are decoded from the file's DataMashup and their M code is mapped onto Applied Steps. See [docs/GRADING.md](docs/GRADING.md) for the details.

## Adding or changing content

Course content is plain JavaScript in `assets/js/content/`. Copy `assets/js/content/quiz-template.js` to start a new quiz. It shows every question type with comments. [docs/AUTHORING.md](docs/AUTHORING.md) is the full reference for lessons, workbooks and questions. Teachers who prefer not to edit files can use the quiz builder on the **Quiz Templates** page.

## Project layout

```
index.html                  the site (loads the scripts below in order)
assets/css/styles.css       white-and-green theme, responsive, print styles
assets/js/core/             no-DOM engines, also usable from Node
  formula.js                Excel formula parser and evaluator (96 functions)
  workbook.js               sheets, cells, tables, validation, sorting, undo
  pivot.js                  PivotTable summaries
  query.js                  Power Query: sources and Applied Steps
  xlsx.js                   read and write real .xlsx files (zip, XML, DataMashup)
  grader.js                 answer keys and grading
assets/js/ui/               spreadsheet grid, charts, Power Query editor, DOM helpers
assets/js/content/          course data, 15 lessons, 12 quizzes, 3 exams, template
assets/js/app.js            pages, routing, submissions, gradebook, quiz builder
docs/                       authoring and grading guides
scripts/serve.js            local static server
test/                       node:test suites, fixtures, browser tests
```

## Development

Requirements: Node 22 or newer. Playwright and LibreOffice are optional; they enable the browser tests and the LibreOffice round-trip test.

```bash
npm install                 # installs Playwright for the browser tests
npx playwright install chromium
npm run serve               # http://localhost:8080
npm test                    # every suite
npm run test:unit           # engines only, no browser
npm run test:content        # every answer key scores 100%, blank scores < 10%
npm run test:excel          # .xlsx round trips, Excel fixtures, LibreOffice re-save
npm run test:e2e            # the site in headless Chromium
```

The test suites cover:

- **formula, workbook, pivot, query and grader:** the engines, against Excel's behavior.
- **xlsx:** write/read round trips; a file made by openpyxl; Excel-style PivotTable and Power Query parts; and an uploaded workbook graded end to end.
- **libreoffice:** every answer key written as `.xlsx`, re-saved by LibreOffice, read back and graded. It is skipped when LibreOffice is not installed.
- **content:** the syllabus calendar, lesson completeness, that every formula parses, and every assessment.
- **e2e:** every page loads without script errors, a practice check, a quiz taken through the UI, an Excel upload, the quiz builder, and no sideways scrolling at phone width.

`test/fixtures/excel-structures.xlsx` is generated by `npm run fixtures` (Python with openpyxl).

## Publishing

The site is static, so any web host works. With GitHub Pages, the included workflow (`.github/workflows/pages.yml`) publishes the default branch. To turn it on, open **Settings → Pages** and set **Source** to **GitHub Actions**. CI (`.github/workflows/ci.yml`) runs every test suite on each push and pull request.

## Browser support

Current Chrome, Edge, Firefox and Safari, including Chromebooks. Reading uploaded `.xlsx` files uses the browser's built-in `DecompressionStream`.
