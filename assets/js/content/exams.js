/*
 * exams.js — the three unit exams (Fridays of weeks 5, 10 and 15). Each covers
 * only its own third of the course.
 */
(function (root) {
  'use strict';
  const isNode = typeof module === 'object' && module.exports;
  const C = isNode ? require('./course.js') : root.DIM.Content;
  const D = C.data;
  const N = D.storeSales.length + 1;
  const exam = (e) => C.exams.push({ minutes: 50, kind: 'exam', ...e });

  // ---------------------------------------------------------------- Exam 1 (weeks 1–5)
  exam({
    id: 'exam-1', week: 5, unit: 1, covers: [1, 2, 3, 4, 5], title: 'Exam 1 · Data foundations and Excel basics',
    intro: 'Weeks 1–5. Part A is on this page. Part B is in the workbook: an equipment inventory to fix, measure and organize.',
    tools: 'organize',
    workbook: {
      sheets: [{
        name: 'Equipment',
        rows: [
          ['Tag', 'Item', 'Room', 'Purchased', 'Cost', 'Condition'],
          ["'0041", 'Projector', 'B12', '8/14/2023', 420, 'Good'],
          ["'0042", 'Laptop Cart', 'Library', '1/9/2024', 1650, 'Good'],
          ["'0043", 'Document Camera', 'B12', '8/14/2023', 189, 'Fair'],
          ["'0044", 'Printer', 'Office', '3/2/2022', 310, 'Poor'],
          ["'0045", 'Projector', 'C04', '8/20/2024', 455, 'Good'],
          ["'0046", 'Speaker', 'Gym', null, 95, 'Fair'],
          ["'0047", 'Laptop Cart', 'C04', '1/9/2024', 1650, null],
          ["'0048", 'Scanner', 'Office', '5/5/2021', 140, 'Poor'],
        ],
        cells: { H1: { input: 'Measure', style: { bold: true }, locked: true }, I1: { input: 'Answer', style: { bold: true }, locked: true }, H2: { input: 'Total cost', locked: true }, H3: { input: 'Average cost (rounded to 2 decimals)', locked: true }, H4: { input: 'Items with a purchase date (COUNT)', locked: true }, H5: { input: 'Condition completeness (%)', locked: true } },
        formats: { 'D2:D9': 'date', 'E2:E9': 'currency', I2: 'currency', I3: 'currency', I5: 'percent' },
        widths: { B: 140, H: 260 },
        styles: { 'I2:I5': { fill: 'answer' } },
      }],
    },
    questions: [
      { id: 'q1', type: 'choice', points: 2, prompt: '"The library\'s laptop cart was checked out 41 times in March, more than any other item" is…', options: ['Data', 'Information', 'Knowledge', 'Unstructured data'], answer: 1 },
      { id: 'q2', type: 'match', points: 4, prompt: 'Classify each field.', pairs: [['Room number (B12, C04)', 'Nominal'], ['Condition (Poor, Fair, Good)', 'Ordinal'], ['Number of times checked out', 'Discrete'], ['Cost in dollars', 'Continuous']] },
      { id: 'q3', type: 'choice', points: 2, prompt: 'Photos from a security camera are…', options: ['Structured data', 'Semi-structured data', 'Unstructured data', 'Not data'], answer: 2 },
      { id: 'q4', type: 'choice', points: 2, prompt: 'Why is 8/14/2023 right-aligned when typed in a cell?', options: ['It is text', 'Excel stores dates as numbers', 'It is a formula', 'Dates are always centered'], answer: 1 },
      { id: 'q5', type: 'match', points: 4, prompt: 'Match the problem to the data quality dimension.', pairs: [['Speaker has no purchase date', 'Completeness'], ['"Laptop Cart" and "laptop cart"', 'Consistency'], ['A cost of -95', 'Validity'], ['The same tag listed twice', 'Uniqueness']] },
      { id: 'q6', type: 'choice', points: 2, prompt: 'Only students who finished the course fill out the course survey, so it looks better than it was. This is…', options: ['Response bias', 'Survivorship bias', 'Timeliness', 'A primary source'], answer: 1 },
      { id: 'q7', type: 'choice', points: 2, prompt: 'You fill =E2/E10 down from F2. What will F3 contain?', options: ['=E3/E10', '=E3/E11', '=E2/E11', '=E2/E10'], answer: 1 },
      { id: 'q8', type: 'choice', points: 2, prompt: 'A coworker needs your data imported into the attendance system. Send…', options: ['.xlsx with formatting', '.csv', 'A screenshot', 'A PivotTable'], answer: 1 },
      { id: 'q9', type: 'choice', points: 2, prompt: 'Opening a CSV of zip codes by double-clicking turns 02134 into 2134. The best fix is…', options: ['Retype each zip', 'Import with Data › From Text/CSV and set Zip to Text', 'Format the cells as Currency', 'Use SUM'], answer: 1 },
      { id: 'q10', type: 'choice', points: 2, prompt: 'Which prevents a bad value from being typed in the first place?', options: ['Conditional formatting', 'Sorting', 'Data validation', 'Filtering'], answer: 2 },
      { id: 'q11', type: 'cellType', sheet: 'Equipment', range: 'A2:A9', valueType: 'text', checkValue: false, points: 2, prompt: 'Part B: the Tag column must stay text (do not break it).' },
      { id: 'q12', type: 'cell', sheet: 'Equipment', cell: 'I2', solution: '=SUM(E2:E9)', functions: ['SUM'], points: 3, prompt: 'I2: total cost with SUM.' },
      { id: 'q13', type: 'cell', sheet: 'Equipment', cell: 'I3', solution: '=ROUND(AVERAGE(E2:E9),2)', functions: ['ROUND', 'AVERAGE'], points: 3, prompt: 'I3: average cost rounded to 2 decimals.' },
      { id: 'q14', type: 'cell', sheet: 'Equipment', cell: 'I4', solution: '=COUNT(D2:D9)', functions: ['COUNT'], points: 3, prompt: 'I4: items with a purchase date.' },
      { id: 'q15', type: 'cell', sheet: 'Equipment', cell: 'I5', solution: '=COUNTA(F2:F9)/COUNTA(A2:A9)', requireFormula: true, tolerance: 0.001, points: 3, prompt: 'I5: completeness of Condition as a percent.' },
      { id: 'q16', type: 'sorted', sheet: 'Equipment', range: 'A1:F9', column: 'E', order: 'desc', points: 3, prompt: 'Sort the inventory by Cost, largest first (keep rows together).' },
      { id: 'q17', type: 'table', sheet: 'Equipment', range: 'A1:F9', name: 'Equipment', points: 3, prompt: 'Format A1:F9 as a Table named Equipment.' },
      { id: 'q18', type: 'validation', sheet: 'Equipment', range: 'F2:F9', solution: { type: 'list', values: ['Good', 'Fair', 'Poor'] }, points: 3, prompt: 'List validation on Condition (Good, Fair, Poor).' },
      { id: 'q19', type: 'condFormat', sheet: 'Equipment', range: 'E2:E9', solution: { op: '>', value: 1000, style: 'red' }, points: 2, prompt: 'Highlight costs greater than 1000.' },
    ],
  });

  // ---------------------------------------------------------------- Exam 2 (weeks 6–10)
  exam({
    id: 'exam-2', week: 10, unit: 2, covers: [6, 7, 8, 9, 10], title: 'Exam 2 · Analyzing and visualizing data',
    intro: 'Weeks 6–10. Part A is on this page. Part B uses the store data: lookups, conditional functions, cleaning, a pivot and a chart.',
    tools: 'analyze',
    workbook: {
      sheets: [
        (() => {
          const s = D.storeSalesSheet('Sales', { extraHeaders: ['Revenue', 'Supplier', 'Code'], formats: { ['J2:J' + N]: 'currency' }, widths: { J: 80, K: 120, L: 70 } });
          s.cells = {};
          for (let r = 2; r <= N; r++) s.cells['J' + r] = { input: '=F' + r + '*G' + r, locked: true };
          s.lock = ['A1:J' + N, 'K1:L1'];
          s.styles = { ['K2:L' + N]: { fill: 'answer' } };
          return s;
        })(),
        { name: 'Suppliers', rows: [['ProductID', 'Supplier'], ...D.catalog.map((p) => [p[0], p[2] === 'Apparel' ? 'Spirit Wear Co' : p[2] === 'Supplies' ? 'OfficeHub' : 'SnackWorks'])], lock: true, widths: { B: 120 } },
        { name: 'Answers', rows: [['Question', 'Answer'], ['Drinks revenue paid by Card'], ['Orders with Qty of 5 or more'], ['Highest single Snacks order revenue'], ['Revenue supplied by OfficeHub']], formats: { B2: 'currency', B4: 'currency', B5: 'currency' }, lock: 'A1:A5', widths: { A: 300, B: 100 }, styles: { 'B2:B5': { fill: 'answer' } } },
      ],
    },
    questions: [
      { id: 'q1', type: 'choice', points: 2, prompt: 'Which is a KPI?', options: ['The list of cashiers', 'Average order value, with a target of $8', 'The Category column', 'A bar chart'], answer: 1 },
      { id: 'q2', type: 'choice', points: 2, prompt: '=COUNTIFS(E2:E61,"Drinks",H2:H61,"Cash") counts orders that are…', options: ['Drinks or Cash', 'Drinks and Cash', 'Not Drinks', 'All orders'], answer: 1 },
      { id: 'q3', type: 'choice', points: 2, prompt: 'In Enrollments(EnrollmentID, StudentID, CourseID), the foreign keys are…', options: ['EnrollmentID only', 'StudentID and CourseID', 'All three', 'None'], answer: 1 },
      { id: 'q4', type: 'choice', points: 2, prompt: 'Why is a student\'s name a poor primary key?', options: ['Names are too long', 'Names are not guaranteed unique and can change', 'Names are text', 'Names cannot be sorted'], answer: 1 },
      { id: 'q5', type: 'number', points: 2, prompt: 'What does =MID("ROBOTICS-2026", 10, 4) return?', answer: 2026, explain: 'MID returns the text "2026"; typing 2026 is accepted.' },
      { id: 'q6', type: 'match', points: 4, prompt: 'Match each question to the chart.', pairs: [['Monthly revenue over a year', 'Line'], ['Revenue by category', 'Bar'], ['Card vs cash share', 'Pie'], ['Price vs quantity sold', 'Scatter']] },
      { id: 'q7', type: 'choice', points: 2, prompt: 'In a PivotTable, "Sum of OrderID" is a problem because…', options: ['IDs cannot be summed', 'Adding ID numbers is meaningless; count them instead', 'It is too slow', 'Pivots cannot show IDs'], answer: 1 },
      { id: 'q8', type: 'choice', points: 2, prompt: 'Grouping a date field by month in a pivot changes the…', options: ['Aggregation', 'Granularity', 'Data type', 'Source'], answer: 1 },
      { id: 'q9', type: 'range', sheet: 'Sales', range: 'K2:K' + N, solution: '=XLOOKUP(C2,Suppliers!$A$2:$A$13,Suppliers!$B$2:$B$13)', functions: [['XLOOKUP', 'VLOOKUP', 'INDEX']], points: 4, prompt: 'K: each order\'s Supplier, looked up by ProductID.' },
      { id: 'q10', type: 'range', sheet: 'Sales', range: 'L2:L' + N, solution: '=LEFT(E2,3)&"-"&C2', requireFormula: true, points: 3, prompt: 'L: a code like "Sna-P103" (first 3 letters of Category, a dash, then ProductID).' },
      { id: 'q11', type: 'cell', sheet: 'Answers', cell: 'B2', solution: '=SUMIFS(Sales!J2:J61,Sales!E2:E61,"Drinks",Sales!H2:H61,"Card")', tolerance: 0.005, points: 3, prompt: 'Drinks revenue paid by Card.' },
      { id: 'q12', type: 'cell', sheet: 'Answers', cell: 'B3', solution: '=COUNTIF(Sales!F2:F61,">=5")', points: 3, prompt: 'Orders with Qty ≥ 5.' },
      { id: 'q13', type: 'cell', sheet: 'Answers', cell: 'B4', solution: '=MAXIFS(Sales!J2:J61,Sales!E2:E61,"Snacks")', tolerance: 0.005, points: 3, prompt: 'Highest single Snacks order revenue.' },
      { id: 'q14', type: 'cell', sheet: 'Answers', cell: 'B5', solution: '=SUMIF(Sales!K2:K61,"OfficeHub",Sales!J2:J61)', tolerance: 0.005, points: 3, prompt: 'Revenue supplied by OfficeHub (use your Supplier column or a pivot).' },
      { id: 'q15', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:J61', rows: 'Cashier', values: 'Revenue', agg: 'average' }, points: 4, prompt: 'A PivotTable: Average of Revenue by Cashier.' },
      { id: 'q16', type: 'chart', chartType: ['column', 'bar'], solution: { sheet: 'Pivot1', type: 'column', range: 'A1:B5', title: 'Average order value by cashier' }, anyRange: true, requireTitle: true, axisFromZero: true, points: 3, prompt: 'A titled column or bar chart (axis from 0) of your pivot.' },
    ],
  });

  // ---------------------------------------------------------------- Exam 3 (weeks 11–15)
  const fileApr = (() => {
    const lines = ['Order ID,Order Date,Product ID,Qty,Unit Price,Store'];
    const r = C.rng(404);
    const stores = ['Main Office', ' main office', 'Cafeteria', 'CAFETERIA ', 'Gym Kiosk'];
    for (let i = 0; i < 14; i++) {
      const p = D.catalog[Math.floor(r() * D.catalog.length)];
      lines.push([4001 + i, '4/' + (1 + Math.floor(r() * 28)) + '/2026', p[0], 1 + Math.floor(r() * 4), '$' + p[3].toFixed(2), stores[Math.floor(r() * stores.length)]].join(','));
      if (i === 3 || i === 9) lines.push(lines[lines.length - 1]);
      if (i === 6) lines.push(',,,,,');
    }
    lines.push('4099,4/30/2026,P777,2,$3.00,Cafeteria');
    return lines.join('\n');
  })();
  D.files['sales_apr.csv'] = fileApr;

  exam({
    id: 'exam-3', week: 15, unit: 3, covers: [11, 12, 13, 14, 15], title: 'Exam 3 · Power Query, Power BI and data ethics',
    intro: 'Weeks 11–15. Part A is on this page. Part B is a Power Query build on April\'s export.',
    tools: ['query', 'format', 'sheets', 'download', 'reset'],
    workbook: {
      sheets: [{ name: 'Answers', rows: [['Question', 'Answer'], ['Rows in AprilClean'], ['Cafeteria revenue in April'], ['Orders whose product is not in the catalog (left anti)']], formats: { B3: 'currency' }, lock: 'A1:A4', widths: { A: 380, B: 110 }, styles: { 'B2:B4': { fill: 'answer' } } }],
      sources: { 'sales_apr.csv': fileApr, 'products.csv': D.files['products.csv'] },
    },
    questions: [
      { id: 'q1', type: 'choice', points: 2, prompt: 'Which statement about Power Query is true?', options: ['It edits the source file directly', 'It records steps that re-run on refresh, leaving the source unchanged', 'It only works on .xlsx files', 'It replaces formulas'], answer: 1 },
      { id: 'q2', type: 'choice', points: 2, prompt: 'A column of prices arrives as "$2.25" text. Before adding a Revenue column you should…', options: ['Delete it', 'Change its type to Currency/Decimal', 'Sort it', 'Unpivot it'], answer: 1 },
      { id: 'q3', type: 'choice', points: 2, prompt: 'Trim is applied after a Filter step on Store = "Cafeteria". What goes wrong?', options: ['Nothing', 'Rows like "CAFETERIA " are dropped before they are cleaned', 'The query will not load', 'The filter is ignored'], answer: 1 },
      { id: 'q4', type: 'match', points: 4, prompt: 'Match each task to the transformation.', pairs: [['Stack January and February', 'Append'], ['Bring in Category from Products', 'Merge'], ['Turn month columns into rows', 'Unpivot'], ['Total revenue per store', 'Group By']] },
      { id: 'q5', type: 'choice', points: 2, prompt: 'Left outer: 60 rows. Inner: 57 rows. What do the 3 missing rows mean?', options: ['Duplicates', 'Sales whose key has no match in the other table', 'Blank rows', 'Errors in the formula'], answer: 1 },
      { id: 'q6', type: 'choice', points: 2, prompt: 'In a star schema, Products is a…', options: ['Fact table', 'Dimension table', 'Measure', 'Slicer'], answer: 1 },
      { id: 'q7', type: 'choice', points: 2, prompt: 'Which DAX measure counts the orders in Sales?', options: ['COUNTROWS(Sales)', 'SUM(Sales)', 'ROWS(Sales[OrderID])', 'COUNTIF(Sales)'], answer: 0 },
      { id: 'q8', type: 'choice', points: 2, prompt: 'You want "Profit Margin = Profit / Revenue" in a card that responds to slicers. Create a…', options: ['Calculated column', 'Measure using DIVIDE', 'New table', 'Text box'], answer: 1 },
      { id: 'q9', type: 'match', points: 4, prompt: 'Classify each field.', pairs: [['Home address', 'Direct identifier'], ['Zip code + birth date', 'Indirect identifiers'], ['Health condition', 'Sensitive data'], ['Number of club members', 'Aggregate (not personal)']] },
      { id: 'q10', type: 'choice', points: 2, prompt: 'A report about club attendance includes each student\'s name. The best change is…', options: ['Keep names for accuracy', 'Remove the name column in Power Query (data minimization)', 'Make the names smaller', 'Sort by name'], answer: 1 },
      { id: 'q11', type: 'query', solution: { name: 'AprilClean', source: { kind: 'file', file: 'sales_apr.csv' }, steps: [{ type: 'removeBlankRows' }, { type: 'removeDuplicates', columns: [] }, { type: 'transform', column: 'Store', op: 'trim' }, { type: 'transform', column: 'Store', op: 'proper' }, { type: 'changeType', columns: { 'Order Date': 'date', Qty: 'whole', 'Unit Price': 'currency' } }, { type: 'addColumn', name: 'Revenue', formula: '[Qty]*[Unit Price]', dataType: 'currency' }] }, steps: ['removeBlankRows', 'removeDuplicates', 'transform', 'changeType', 'addColumn'], loaded: true, points: 10, prompt: 'Part B: AprilClean query on sales_apr.csv: blank rows, duplicates, trim and capitalize Store, types, Revenue, loaded.' },
      { id: 'q12', type: 'cell', sheet: 'Answers', cell: 'B2', solution: '=COUNTA(AprilClean!A:A)-1', points: 3, prompt: 'Rows in AprilClean.' },
      { id: 'q13', type: 'cell', sheet: 'Answers', cell: 'B3', solution: '=SUMIF(AprilClean!F:F,"Cafeteria",AprilClean!G:G)', tolerance: 0.005, points: 3, prompt: 'Cafeteria revenue in April.' },
      { id: 'q14', type: 'cell', sheet: 'Answers', cell: 'B4', expect: 1, solution: 1, points: 3, prompt: 'Orders whose product is not in products.csv (try a left anti merge).' },
    ],
  });

  if (isNode) module.exports = C;
})(typeof globalThis !== 'undefined' ? globalThis : this);
