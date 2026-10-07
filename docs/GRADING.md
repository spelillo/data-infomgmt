# How grading works

Every graded item on the site is graded automatically when it is submitted. Practice workbooks use the same grader through **Check my work**. This page explains what is compared, how partial credit works, how uploads from Excel are read, and how the teacher's gradebook verifies what students hand in.

## The answer key

There is no hand-typed answer sheet. For each assessment the grader builds a key workbook:

1. It starts from the assessment's starting workbook.
2. It runs the author's optional `model` walkthrough (for example "import this file, then remove duplicates").
3. It applies every question's `solution`: a formula, a value, a table, a rule, a chart, a PivotTable or a query.

Because solutions are usually formulas, the key is computed from the same data the student sees. The content test suite builds every key and checks two things: the key scores exactly 100%, and a blank submission scores under 10%.

## What is compared

- **Cells and ranges.** The student's value is compared to the key's value at the same address. Numbers match within a tiny rounding tolerance, or the question's `tolerance`. Text ignores case and extra spaces unless the question says otherwise. Errors must be the same error.
- **Method.** When a task asks for a formula, a correct value that was typed in gets half credit. A formula missing a required function or reference loses a quarter. The feedback says which rule applied.
- **How values are stored.** `cellType` questions check that IDs with leading zeros are text and that dates are real dates.
- **Structures.** Tables, data validation, conditional formatting, sorting, charts, PivotTables and Power Query queries are checked by what they are and what they produce, not by where they appear on screen. For example, a query is checked for its step types and the number of rows it outputs. A sort is checked to confirm whole records stayed together.

Feedback is shown next to every item. Before submission, practice can reveal expected values with **Show expected answers**. After a graded submission, feedback includes the expected value for anything that was not correct.

## Uploading from Excel

Students on desktop Excel download the blank template, work in Excel, and upload the `.xlsx`. The upload is graded by exactly the same rules as the browser workbook. The reader (`assets/js/core/xlsx.js`) understands:

| In the file | Read as |
|---|---|
| Values, formulas, cached results | Cells. Formulas are recalculated, and shared formulas from fill-down are expanded. |
| Structured references (`Sales[Qty]`, `Sales[[#This Row],[Qty]]`) | Table references, evaluated against the table. |
| Number formats | Currency, percent, date, text and the others the course uses. |
| Excel Tables (`xl/tables`) | Tables with their name and range. |
| Data validation, including lists that point at a range or another sheet | Validation rules. |
| Conditional formatting (cell value, text contains, duplicates, top N, above average, blanks) | Highlight rules, with the fill read as green, red or yellow. |
| Charts (`xl/charts`) | Type (column, bar, line, pie, scatter, area), data range, title (typed, linked or automatic) and value-axis minimum. |
| PivotTables (`xl/pivotTables`, `xl/pivotCache`) | Source range or table, row, column and value fields, summary function, and date grouping (months, quarters, years). |
| Power Query (`customXml` DataMashup) | Each query's source (file, folder, workbook table or another query) and its M steps, mapped to Applied Steps. The loaded table is linked through the workbook's connections. |

Power Query steps are mapped from M functions. For example, `Table.Distinct` becomes Remove Duplicates, `Table.TransformColumns` with `Text.Trim` becomes Trim, `Table.NestedJoin` becomes Merge (with its join kind), `Table.Group` becomes Group By, and `Table.UnpivotOtherColumns` becomes Unpivot. Imported queries are not re-run. The rows Excel loaded into the sheet are the query's output, so the check is "did your query produce the right table."

A query that is "connection only" (not loaded to a sheet) is reported as such, so the student knows to use **Close & Load To… → Table**.

The site writes cells, number formats, column widths, names, Tables, data validation and conditional formatting into `.xlsx`. Charts, PivotTables and queries built in the browser are not written to the file, because students build those in Excel when they work there.

## Receipts and the gradebook

When a student submits, the browser records the work, the answers and the score, and offers a **receipt** file (`.dim.json`). The receipt contains the workbook itself, any uploaded file under 3 MB, and the answers.

On **Grades**, the teacher imports receipts. Each one is re-graded from the workbook inside it, so the score does not depend on the student's browser. If the re-graded score differs from the score in the receipt, the row is flagged. The teacher scores rubric points (`manualPoints`) on the same page and exports the gradebook as CSV.

**Settings** decides whether a student's first or best attempt counts.

## The course grade

| Category | Weight | Notes |
|---|---|---|
| Assignments | 30% | Average of submitted assignments. |
| Quizzes | 15% | The lowest quiz is dropped once there are two or more. |
| Exam 1, 2, 3 | 15% each | |
| Project | 10% | The week-14 dashboard. |

Categories with nothing submitted yet are left out, and the remaining weights are rescaled, so a running grade is meaningful from week one.

## Tests that guard grading

- `test/grader.test.js` covers partial credit, method rules, feedback and every question type.
- `test/content.test.js` checks that every assessment's key scores 100% and a blank submission scores under 10%.
- `test/xlsx.test.js` covers round trips, a workbook from another spreadsheet program, Excel-style PivotTable and Power Query parts, and an upload graded end to end.
- `test/libreoffice.test.js` writes every key as `.xlsx`, re-saves it with LibreOffice, reads it back and grades it at 100%.
- `test/e2e.test.js` takes a quiz and uploads an assignment through the real site.
