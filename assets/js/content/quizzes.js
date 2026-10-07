/*
 * quizzes.js — the twelve Monday quizzes. Each covers the previous week only.
 * No quiz in week 1 or on the Monday after an exam (weeks 6 and 11).
 *
 * A quiz is an assessment: { id, title, week (when it is given), covers,
 * minutes, intro, workbook?, questions }. Workbook questions are graded from
 * the quiz workbook (web grid or uploaded .xlsx); the rest from the page.
 */
(function (root) {
  'use strict';
  const isNode = typeof module === 'object' && module.exports;
  const C = isNode ? require('./course.js') : root.DIM.Content;
  const D = C.data;
  const N = D.storeSales.length + 1;

  const quiz = (q) => C.quizzes.push({ minutes: 15, kind: 'quiz', ...q });

  // ---------------------------------------------------------------- Quiz 2 (covers week 1)
  quiz({
    id: 'quiz-2', week: 2, covers: 1, title: 'Quiz 2 · What data is',
    intro: 'Five page questions and two workbook tasks. Use the status bar and the Name Box; no formulas needed.',
    tools: 'basic',
    workbook: { sheets: [{ name: 'Survey', rows: D.survey.map((r, i) => (i === 0 ? [...r, null, 'Task', 'Answer'] : r)), cells: { H2: 'Total minutes to school for all 15 students (status bar Sum of F2:F16)', H3: 'Number of students in grade 11' }, lock: true, widths: { C: 110, D: 105, H: 380 }, styles: { 'I2:I3': { fill: 'answer' } } }] },
    questions: [
      { id: 'q1', type: 'choice', prompt: '"5 of 15 students chose pizza" is an example of…', options: ['Data', 'Information', 'Knowledge', 'A decision'], answer: 1, explain: 'Counting and summarizing raw answers turns data into information.' },
      { id: 'q2', type: 'choice', prompt: 'Which is raw data?', options: ['Pizza is the most popular lunch', 'The cafeteria should order more pizza', '"Tacos" typed in cell C3', 'Lunch sales went up 20%'], answer: 2 },
      { id: 'q3', type: 'match', prompt: 'Match each term to its meaning.', pairs: [['Cell address', 'Column letter and row number, like C4'], ['Range', 'A block of cells such as A2:F16'], ['Worksheet', 'One grid of cells in a workbook'], ['Formula bar', 'Shows what is really stored in the active cell']] },
      { id: 'q4', type: 'choice', prompt: 'In a well-organized table, each row should be…', options: ['One kind of fact', 'One record (one thing being described)', 'A total', 'A different table'], answer: 1 },
      { id: 'q5', type: 'text', prompt: 'Which key confirms an entry and moves one cell to the right?', accept: ['Tab', 'Tab key'] },
      { id: 'q6', type: 'cell', sheet: 'Survey', cell: 'I2', solution: '=SUM(F2:F16)', points: 2, prompt: 'I2: total minutes to school (read the status bar Sum).' },
      { id: 'q7', type: 'cell', sheet: 'Survey', cell: 'I3', solution: '=COUNTIF(B2:B16,11)', points: 2, prompt: 'I3: how many students are in grade 11.' },
    ],
  });

  // ---------------------------------------------------------------- Quiz 3 (covers week 2)
  quiz({
    id: 'quiz-3', week: 3, covers: 2, title: 'Quiz 3 · Data types and tidy tables',
    intro: 'Classify, fix types, and spot untidy data.',
    tools: 'basic',
    workbook: {
      sheets: [{
        name: 'Fix',
        rows: [['Item', 'Price', 'Sold On', 'Product Code'], ['Notebook', "'3.50", "'2/2/2026", '501'], ['Hoodie', "'32", "'2/3/2026", '502'], ['Lanyard', "'5", "'2/5/2026", '507']],
        lock: ['A1:D1', 'A2:A4'],
        styles: { 'B2:D4': { fill: 'answer' } },
      }],
    },
    questions: [
      { id: 'q1', type: 'choice', prompt: 'T-shirt size (S, M, L, XL) is…', options: ['Nominal', 'Ordinal', 'Discrete', 'Continuous'], answer: 1 },
      { id: 'q2', type: 'choice', prompt: 'A zip code like 02134 should be stored as…', options: ['A number, because it is digits', 'Text, because you never do math on it and the leading zero matters', 'A date', 'A percent'], answer: 1 },
      { id: 'q3', type: 'number', prompt: 'Excel stores 1/1/1900 as 1. What number is 1/3/1900?', answer: 3 },
      { id: 'q4', type: 'choice', multi: true, prompt: 'Which break the tidy-data rules? (select all)', options: ['A column called "Jan" and another called "Feb" holding the same measure', 'A cell containing "15 min"', 'One header row with unique names', 'A Total row inside the data'], answer: [0, 1, 3] },
      { id: 'q5', type: 'choice', prompt: 'You format 0.256 as Percent with no decimals. What is stored in the cell?', options: ['26%', '0.26', '0.256', '26'], answer: 2 },
      { id: 'q6', type: 'cellType', sheet: 'Fix', range: 'B2:B4', valueType: 'number', expect: [3.5, 32, 5], points: 2, prompt: 'Make Price (B2:B4) real numbers.' },
      { id: 'q7', type: 'cellType', sheet: 'Fix', range: 'C2:C4', valueType: 'date', expect: [46055, 46056, 46058], points: 2, prompt: 'Make Sold On (C2:C4) real dates.' },
      { id: 'q8', type: 'cellType', sheet: 'Fix', range: 'D2:D4', valueType: 'text', expect: ['00501', '00502', '00507'], points: 2, prompt: 'Product codes must be 5-character text with leading zeros (00501…).' },
    ],
  });

  // ---------------------------------------------------------------- Quiz 4 (covers week 3)
  quiz({
    id: 'quiz-4', week: 4, covers: 3, title: 'Quiz 4 · Data quality and formulas',
    intro: 'Quality dimensions, bias, and your first functions.',
    tools: 'basic',
    workbook: {
      sheets: [{
        name: 'Scores',
        rows: [['Student', 'Score', 'Bonus', 'Final', null, 'Measure', 'Answer'], ['Maya', 88, 2], ['Luis', null, 5], ['Aisha', 93, 0], ['Ethan', 71, 3], ['Grace', 85, null], ['Noah', null, 1]].map((r, i) => (i === 0 ? r : [r[0], r[1], r[2]])),
        cells: { F2: 'Scores recorded (COUNT)', F3: 'Missing scores (COUNTBLANK)', F4: 'Average score, rounded to 0 decimals', H1: { input: 'Bonus multiplier', style: { bold: true } }, H2: 1.5 },
        lock: ['A1:G1', 'A2:C7', 'F2:F4', 'H1:H2'],
        widths: { F: 250 },
        styles: { 'D2:D7': { fill: 'answer' }, 'G2:G4': { fill: 'answer' } },
      }],
    },
    questions: [
      { id: 'q1', type: 'match', prompt: 'Match each problem to its quality dimension.', pairs: [['"Robotics" and "robotics club" in one column', 'Consistency'], ['A grade of 15 in a high school', 'Validity'], ['The same sign-up appears twice', 'Uniqueness'], ['Last year\'s roster used for this year\'s buses', 'Timeliness']] },
      { id: 'q2', type: 'choice', prompt: 'A survey about study habits is handed out only in the library. This is…', options: ['Response bias', 'Selection (sampling) bias', 'Survivorship bias', 'Not biased'], answer: 1 },
      { id: 'q3', type: 'choice', prompt: 'What does =COUNTA(A2:A7) count?', options: ['Cells with numbers', 'Cells that are not empty', 'Empty cells', 'Cells with text only'], answer: 1 },
      { id: 'q4', type: 'number', prompt: 'What does =2+3*4^2 return?', answer: 50 },
      { id: 'q5', type: 'range', sheet: 'Scores', range: 'D2:D7', solution: '=B2+C2*$H$2', requireFormula: true, mustReference: ['$H$2'], points: 3, prompt: 'D2:D7: Final = Score + Bonus × the multiplier in H2 (absolute reference), filled down.' },
      { id: 'q6', type: 'cell', sheet: 'Scores', cell: 'G2', solution: '=COUNT(B2:B7)', functions: ['COUNT'], prompt: 'G2: scores recorded.' },
      { id: 'q7', type: 'cell', sheet: 'Scores', cell: 'G3', solution: '=COUNTBLANK(B2:B7)', functions: ['COUNTBLANK'], prompt: 'G3: missing scores.' },
      { id: 'q8', type: 'cell', sheet: 'Scores', cell: 'G4', solution: '=ROUND(AVERAGE(B2:B7),0)', functions: ['ROUND', 'AVERAGE'], prompt: 'G4: average score rounded to a whole number.' },
    ],
  });

  // ---------------------------------------------------------------- Quiz 5 (covers week 4)
  quiz({
    id: 'quiz-5', week: 5, covers: 4, title: 'Quiz 5 · CSV vs Excel and delimiters',
    intro: 'Import a semicolon file without losing data, then answer the format questions.',
    tools: ['format', 'import', 'textToColumns', 'csv', 'sheets', 'download', 'reset'],
    workbook: {
      sheets: [{ name: 'Start Here', rows: [['Import products_eu.csv (semicolon-delimited). Keep Code as text. Rename the sheet to Products.']], headerRow: false, lock: true, widths: { A: 560 } }],
      sources: { 'products_eu.csv': 'Code;Product;Stock\n0101;Granola Bar;40\n0102;Trail Mix;25\n0201;Water Bottle;60\n0301;Notebook;18' },
    },
    model: (wb) => wb.importDelimited(wb.sources['products_eu.csv'], { delimiter: ';', textColumns: [0], sheetName: 'Products' }),
    questions: [
      { id: 'q1', type: 'cellType', sheet: 'Products', range: 'A2:A5', valueType: 'text', expect: ['0101', '0102', '0201', '0301'], points: 3, prompt: 'Products sheet: Code kept as text with leading zeros.' },
      { id: 'q2', type: 'cellType', sheet: 'Products', range: 'C2:C5', valueType: 'number', expect: [40, 25, 60, 18], points: 2, prompt: 'Products sheet: Stock in column C as numbers (split on semicolons).' },
      { id: 'q3', type: 'choice', prompt: 'In the line  00417,"Patel, Maya",02134  the double quotes are a…', options: ['Delimiter', 'Text qualifier', 'Header', 'Formula'], answer: 1 },
      { id: 'q4', type: 'choice', multi: true, prompt: 'Which are lost when a workbook is saved as CSV? (select all)', options: ['Formulas (only results are kept)', 'Extra sheets', 'Cell colors and charts', 'The values in the active sheet'], answer: [0, 1, 2] },
      { id: 'q5', type: 'choice', prompt: 'You must upload a class roster to the district\'s system. The best format is…', options: ['.xlsx with charts', '.csv', '.pdf', '.docx'], answer: 1 },
      { id: 'q6', type: 'text', prompt: 'Name the Excel feature that splits "Patel, Maya" into two columns.', accept: ['Text to Columns', 'Text-to-Columns', 'TextToColumns'] },
    ],
  });

  // ---------------------------------------------------------------- Quiz 7 (covers week 6)
  quiz({
    id: 'quiz-7', week: 7, covers: 6, title: 'Quiz 7 · Questions, metrics and conditional functions',
    intro: 'The Revenue column is filled in. Answer with one formula per cell.',
    tools: 'basic',
    workbook: {
      sheets: [(() => {
        const s = D.storeSalesSheet('Sales', { extraHeaders: ['Revenue'], formats: { ['J2:J' + N]: 'currency', M2: 'currency', M4: 'currency' }, widths: { J: 80, L: 300, M: 90 } });
        s.cells = { L1: { input: 'Question', style: { bold: true } }, M1: { input: 'Answer', style: { bold: true } }, L2: 'Revenue from Supplies', L3: 'Orders taken by Riley', L4: 'Average Card order revenue', L5: 'Label for order 1 (N2): "Big" if Revenue ≥ 15 else "Normal"' };
        for (let r = 2; r <= N; r++) s.cells['J' + r] = { input: '=F' + r + '*G' + r };
        s.lock = ['A1:J' + N, 'L1:M1', 'L2:L5'];
        s.styles = { 'M2:M4': { fill: 'answer' }, N2: { fill: 'answer' } };
        return s;
      })()],
    },
    questions: [
      { id: 'q1', type: 'choice', prompt: 'Which is a measurable question?', options: ['Is the store doing well?', 'What was total Snacks revenue per week in March?', 'Do people like the store?', 'Should we sell more stuff?'], answer: 1 },
      { id: 'q2', type: 'choice', prompt: 'In the store data, Category is a…', options: ['Measure', 'Dimension', 'KPI', 'Formula'], answer: 1 },
      { id: 'q3', type: 'choice', prompt: 'In SUMIFS, the range being added up comes…', options: ['First', 'Last', 'Second', 'Anywhere'], answer: 0 },
      { id: 'q4', type: 'cell', sheet: 'Sales', cell: 'M2', solution: '=SUMIF(E2:E61,"Supplies",J2:J61)', functions: [['SUMIF', 'SUMIFS']], points: 2, prompt: 'M2: revenue from Supplies.' },
      { id: 'q5', type: 'cell', sheet: 'Sales', cell: 'M3', solution: '=COUNTIF(I2:I61,"Riley")', functions: [['COUNTIF', 'COUNTIFS']], points: 2, prompt: 'M3: orders taken by Riley.' },
      { id: 'q6', type: 'cell', sheet: 'Sales', cell: 'M4', solution: '=AVERAGEIF(H2:H61,"Card",J2:J61)', functions: [['AVERAGEIF', 'AVERAGEIFS']], tolerance: 0.005, points: 2, prompt: 'M4: average revenue of Card orders.' },
      { id: 'q7', type: 'cell', sheet: 'Sales', cell: 'N2', solution: '=IF(J2>=15,"Big","Normal")', functions: ['IF'], points: 2, prompt: 'N2: label order 1 with IF.' },
    ],
  });

  // ---------------------------------------------------------------- Quiz 8 (covers week 7)
  quiz({
    id: 'quiz-8', week: 8, covers: 7, title: 'Quiz 8 · Keys and lookups',
    intro: 'Primary keys, foreign keys and XLOOKUP.',
    tools: 'basic',
    workbook: {
      sheets: [
        { name: 'Grades', rows: [['EnrollmentID', 'StudentID', 'Score', 'Advisor'], ['E01', 'S1001', 94], ['E02', 'S1004', 67], ['E03', 'S1009', 98], ['E04', 'S1012', 84], ['E05', 'S1005', 95]], lock: 'A1:D6', styles: { 'D2:D6': { fill: 'answer' } }, widths: { D: 110, F: 220 }, cells: { F1: { input: 'Orphan enrollments', style: { bold: true }, locked: true } } },
        { name: 'Students', rows: D.students, lock: true },
      ],
    },
    questions: [
      { id: 'q1', type: 'choice', prompt: 'Which column is the best primary key for a Students table?', options: ['Last name', 'StudentID', 'Grade', 'Advisor'], answer: 1, explain: 'It is unique, never blank, and does not change.' },
      { id: 'q2', type: 'choice', prompt: 'A foreign key is…', options: ['A key from another country', 'A column that holds another table\'s primary key', 'Any column with unique values', 'The first column of a table'], answer: 1 },
      { id: 'q3', type: 'choice', prompt: 'One advisor has many students. The relationship from Advisors to Students is…', options: ['One-to-one', 'One-to-many', 'Many-to-one', 'Many-to-many'], answer: 1 },
      { id: 'q4', type: 'choice', prompt: 'Why should VLOOKUP end with FALSE for IDs?', options: ['It makes it faster', 'It forces an exact match; otherwise it may return a wrong row', 'It hides errors', 'It sorts the data'], answer: 1 },
      { id: 'q5', type: 'range', sheet: 'Grades', range: 'D2:D6', solution: '=XLOOKUP(B2,Students!$A$2:$A$11,Students!$E$2:$E$11,"Not found")', functions: [['XLOOKUP', 'VLOOKUP', 'INDEX']], points: 3, prompt: 'D2:D6: each student\'s Advisor, "Not found" for orphans.' },
      { id: 'q6', type: 'cell', sheet: 'Grades', cell: 'G1', solution: '=COUNTIF(D2:D6,"Not found")', requireFormula: true, points: 2, prompt: 'G1: count the orphan enrollments.' },
      { id: 'q7', type: 'text', prompt: 'What is a row whose foreign key matches nothing in the related table called? (two words)', accept: ['orphan record', 'orphan records', 'orphan row', 'orphan'] },
    ],
  });

  // ---------------------------------------------------------------- Quiz 9 (covers week 8)
  quiz({
    id: 'quiz-9', week: 9, covers: 8, title: 'Quiz 9 · Cleaning text and dates',
    intro: 'Write one formula in row 2 of each column and fill it down.',
    tools: 'data',
    workbook: {
      sheets: [{
        name: 'Clean',
        rows: [['Raw Name', 'Code', 'Date', 'Name', 'Prefix', 'Month #', 'Label'], ['  zoe MARTIN ', 'snk-101', '3/2/2026'], ['omar haddad', 'drk-202', '3/14/2026'], ['LILY   chen', 'apl-401', '4/1/2026']],
        formats: { 'C2:C4': 'date' },
        lock: ['A1:G1', 'A2:C4'],
        widths: { A: 120, D: 110, G: 140 },
        styles: { 'D2:G4': { fill: 'answer' } },
      }],
    },
    questions: [
      { id: 'q1', type: 'range', sheet: 'Clean', range: 'D2:D4', solution: '=PROPER(TRIM(A2))', functions: ['TRIM', 'PROPER'], points: 2, prompt: 'D: clean Proper Case name.' },
      { id: 'q2', type: 'range', sheet: 'Clean', range: 'E2:E4', solution: '=UPPER(LEFT(B2,3))', functions: ['LEFT', 'UPPER'], points: 2, prompt: 'E: first three letters of Code in capitals (SNK).' },
      { id: 'q3', type: 'range', sheet: 'Clean', range: 'F2:F4', solution: '=MONTH(C2)', functions: ['MONTH'], points: 2, prompt: 'F: month number of the date.' },
      { id: 'q4', type: 'range', sheet: 'Clean', range: 'G2:G4', solution: '=D2&" ("&E2&")"', requireFormula: true, points: 2, prompt: 'G: join name and prefix like "Zoe Martin (SNK)".' },
      { id: 'q5', type: 'number', prompt: 'What does =LEN(TRIM("  data  ")) return?', answer: 4 },
      { id: 'q6', type: 'choice', prompt: 'You will receive this file again every month. What is the best long-term approach?', options: ['Retype the data each month', 'Clean it by hand each month', 'Record the cleaning steps once so they can be repeated (Power Query)', 'Ignore the messy rows'], answer: 2 },
    ],
  });

  // ---------------------------------------------------------------- Quiz 10 (covers week 9)
  quiz({
    id: 'quiz-10', week: 10, covers: 9, title: 'Quiz 10 · Charts that tell the truth',
    intro: 'Pick chart types and spot misleading charts. One chart to build.',
    tools: ['chart', 'sort', 'download', 'reset'],
    workbook: { sheets: [{ name: 'Clubs', rows: [['Club', 'Members'], ['Robotics', 24], ['Debate', 15], ['Art', 31], ['Coding', 19]], lock: 'A1:B1' }] },
    questions: [
      { id: 'q1', type: 'match', prompt: 'Match each question to the best chart.', pairs: [['How did attendance change each month?', 'Line chart'], ['Which club has the most members?', 'Column or bar chart'], ['Do students who sleep more score higher?', 'Scatter plot'], ['What share of orders are card vs cash?', 'Pie chart']] },
      { id: 'q2', type: 'choice', prompt: 'A column chart\'s axis starts at 90. Scores are 91 and 95. The chart…', options: ['Is accurate', 'Exaggerates the difference', 'Hides the difference', 'Cannot be drawn'], answer: 1 },
      { id: 'q3', type: 'choice', multi: true, prompt: 'Which are ways a chart can mislead? (select all)', options: ['A truncated value axis on a bar chart', 'Showing only the months that went up', 'A title that states the finding', 'Two different y-axis scales on one chart'], answer: [0, 1, 3] },
      { id: 'q4', type: 'choice', prompt: 'Best title for a chart of revenue by category?', options: ['Chart 1', 'Revenue by Category', 'Apparel brings in the most revenue', 'Data'], answer: 2 },
      { id: 'q5', type: 'sorted', sheet: 'Clubs', range: 'A1:B5', column: 'B', order: 'desc', points: 2, prompt: 'Sort the clubs by Members, largest first.' },
      { id: 'q6', type: 'chart', sheet: 'Clubs', chartType: ['column', 'bar'], solution: { type: 'bar', range: 'A1:B5', title: 'Art is the largest club' }, requireTitle: true, axisFromZero: true, points: 3, prompt: 'A titled bar or column chart of A1:B5, axis from 0.' },
    ],
  });

  // ---------------------------------------------------------------- Quiz 12 (covers week 11)
  quiz({
    id: 'quiz-12', week: 12, covers: 11, title: 'Quiz 12 · Power Query basics',
    intro: 'Build a short query and answer questions about applied steps.',
    tools: ['query', 'sheets', 'download', 'reset'],
    workbook: { sheets: [{ name: 'Start Here', rows: [['Make a query named Feb from sales_feb.csv: remove blank rows, remove duplicates, trim the Store column, and Close & Load.']], headerRow: false, lock: true, widths: { A: 640 } }], sources: { 'sales_feb.csv': D.files['sales_feb.csv'] } },
    questions: [
      { id: 'q1', type: 'query', solution: { name: 'Feb', source: { kind: 'file', file: 'sales_feb.csv' }, steps: [{ type: 'removeBlankRows' }, { type: 'removeDuplicates', columns: [] }, { type: 'transform', column: 'Store', op: 'trim' }] }, steps: ['removeBlankRows', 'removeDuplicates', 'transform'], loaded: true, points: 4, prompt: 'Feb query with the three steps, loaded.' },
      { id: 'q2', type: 'choice', prompt: 'ETL stands for…', options: ['Enter, Total, List', 'Extract, Transform, Load', 'Edit, Test, Launch', 'Excel, Tables, Lookups'], answer: 1 },
      { id: 'q3', type: 'choice', prompt: 'The Applied Steps list is useful because…', options: ['It makes the file smaller', 'It is a record of every change that can be re-run and audited', 'It hides the data', 'It replaces the source file'], answer: 1 },
      { id: 'q4', type: 'choice', prompt: 'A Product ID column like 0101 should be set to which data type?', options: ['Whole Number', 'Text', 'Date', 'Currency'], answer: 1 },
      { id: 'q5', type: 'choice', prompt: 'March\'s file arrives with the same layout. Fastest correct way to clean it?', options: ['Rebuild every step', 'Point the query at the new file and Refresh', 'Copy and paste the rows into the old sheet', 'Clean it with formulas'], answer: 1 },
    ],
  });

  // ---------------------------------------------------------------- Quiz 13 (covers week 12)
  quiz({
    id: 'quiz-13', week: 13, covers: 12, title: 'Quiz 13 · Combining data',
    intro: 'Joins, appends and unpivot.',
    tools: ['query', 'sheets', 'download', 'reset'],
    workbook: { sheets: [{ name: 'Start Here', rows: [['Make a query named Long from attendance_wide.csv that unpivots the month columns (keep Club). Close & Load.']], headerRow: false, lock: true, widths: { A: 640 } }], sources: { 'attendance_wide.csv': D.files['attendance_wide.csv'] } },
    questions: [
      { id: 'q1', type: 'query', solution: { name: 'Long', source: { kind: 'file', file: 'attendance_wide.csv' }, steps: [{ type: 'unpivot', keep: ['Club'] }] }, steps: ['unpivot'], loaded: true, points: 3, prompt: 'Long query unpivoted and loaded.' },
      { id: 'q2', type: 'choice', prompt: 'January, February and March sales have identical columns. To combine them use…', options: ['Merge', 'Append', 'Unpivot', 'Group By'], answer: 1 },
      { id: 'q3', type: 'choice', prompt: 'Sales has 54 rows. 2 have a Product ID missing from Products. An inner join returns…', options: ['54 rows', '52 rows', '56 rows', '2 rows'], answer: 1 },
      { id: 'q4', type: 'choice', prompt: 'Which join lists only the sales whose product is NOT in the catalog?', options: ['Left outer', 'Inner', 'Left anti', 'Full outer'], answer: 2 },
      { id: 'q5', type: 'number', prompt: 'A wide table has 3 clubs × 6 month columns, with no blanks. How many rows after unpivoting?', answer: 18 },
      { id: 'q6', type: 'choice', prompt: 'Merging Sales with Products is most like which Excel function?', options: ['SUMIF', 'XLOOKUP', 'TRIM', 'COUNTA'], answer: 1 },
    ],
  });

  // ---------------------------------------------------------------- Quiz 14 (covers week 13)
  quiz({
    id: 'quiz-14', week: 14, covers: 13, title: 'Quiz 14 · Power BI reports',
    intro: 'Model and visual questions only.',
    workbook: null,
    questions: [
      { id: 'q1', type: 'choice', prompt: 'In a star schema, the table in the middle holding transactions is the…', options: ['Dimension table', 'Fact table', 'Lookup table', 'Pivot table'], answer: 1 },
      { id: 'q2', type: 'choice', prompt: 'Which Power BI view shows the relationships between tables?', options: ['Report view', 'Table view', 'Model view', 'Query view'], answer: 2 },
      { id: 'q3', type: 'choice', prompt: 'A Store slicer filters Sales because…', options: ['Slicers filter every table equally', 'Filters flow from the "1" side of a relationship to the "many" side', 'Sales is sorted by Store', 'Power BI guesses'], answer: 1 },
      { id: 'q4', type: 'match', prompt: 'Match each visual to its purpose.', pairs: [['Card', 'A single headline number'], ['Matrix', 'A grid like a PivotTable'], ['Slicer', 'Filter buttons for the reader'], ['Line chart', 'A trend over time']] },
      { id: 'q5', type: 'choice', prompt: 'Products is a dimension table. Its relationship to Sales should use…', options: ['Products[ProductID] (1) to Sales[Product ID] (many)', 'Sales[OrderID] (1) to Products[ProductID] (1)', 'Products[Category] to Sales[Store]', 'No relationship; copy the columns with XLOOKUP'], answer: 0 },
      { id: 'q6', type: 'text', prompt: 'What is it called when clicking a bar in one visual filters the other visuals on the page?', accept: ['cross-filtering', 'cross filtering', 'crossfiltering', 'cross-filter', 'cross highlighting', 'cross-highlighting'] },
    ],
  });

  // ---------------------------------------------------------------- Quiz 15 (covers week 14)
  quiz({
    id: 'quiz-15', week: 15, covers: 14, title: 'Quiz 15 · DAX, storytelling and ethics',
    intro: 'DAX measures, dashboard design and personal data.',
    workbook: null,
    questions: [
      { id: 'q1', type: 'choice', prompt: 'Which DAX measure equals =SUMIFS(Revenue, Store, "Cafeteria")?', options: ['CALCULATE([Total Revenue], Sales[Store] = "Cafeteria")', 'SUM(Sales[Store])', 'COUNTROWS(Cafeteria)', 'DIVIDE([Total Revenue], "Cafeteria")'], answer: 0 },
      { id: 'q2', type: 'choice', prompt: 'A measure differs from a calculated column because a measure…', options: ['Is stored once per row', 'Is recalculated for whatever filters the visual shows', 'Can only hold text', 'Cannot be used in a card'], answer: 1 },
      { id: 'q3', type: 'match', prompt: 'Classify each field.', pairs: [['Student email', 'Direct identifier'], ['Birth date', 'Indirect identifier'], ['Free/reduced lunch status', 'Sensitive data'], ['Favorite lunch', 'Not personal data on its own']] },
      { id: 'q4', type: 'choice', prompt: 'A chart shows "Grade 11 Web Design students with an IEP: 1". The best fix is…', options: ['Add the student\'s name for clarity', 'Suppress or combine groups smaller than about 5', 'Make the bar red', 'Leave it, since the number is accurate'], answer: 1 },
      { id: 'q5', type: 'choice', prompt: 'The US law that protects student education records is…', options: ['HIPAA', 'FERPA', 'COPPA', 'GDPR'], answer: 1 },
      { id: 'q6', type: 'choice', prompt: 'Where should the most important number go on a dashboard?', options: ['Bottom-right', 'Top-left', 'In the legend', 'On a second page'], answer: 1 },
      { id: 'q7', type: 'text', prompt: 'Which DAX function divides two numbers and returns BLANK instead of an error when dividing by zero?', accept: ['DIVIDE', 'DIVIDE()'] },
    ],
  });

  if (isNode) module.exports = C;
})(typeof globalThis !== 'undefined' ? globalThis : this);
