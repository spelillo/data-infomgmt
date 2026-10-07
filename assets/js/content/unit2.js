/*
 * unit2.js — Unit 2: Analyzing and visualizing data (weeks 6–10).
 */
(function (root) {
  'use strict';
  const isNode = typeof module === 'object' && module.exports;
  const C = isNode ? require('./course.js') : root.DIM.Content;
  const D = C.data;
  const N = D.storeSales.length + 1; // last data row of the store sales sheet (61)

  // Store sales with a Revenue column already filled (weeks 9–10).
  const salesWithRevenue = (name = 'Sales') => {
    const s = D.storeSalesSheet(name, { extraHeaders: ['Revenue'], formats: { ['J2:J' + N]: 'currency' }, widths: { J: 80 } });
    s.cells = {};
    for (let r = 2; r <= N; r++) s.cells['J' + r] = { input: '=F' + r + '*G' + r, locked: true };
    s.lock = 'A1:J' + N;
    return s;
  };

  // ---------------------------------------------------------------- Week 6
  C.lessons.push({
    week: 6, unit: 2, id: 'week-6',
    title: 'From question to metric',
    summary: 'Turning vague questions into measurable ones, metrics and KPIs, dimensions and measures, and the conditional functions that compute them: IF, IFS, AND/OR, COUNTIFS, SUMIFS and AVERAGEIFS.',
    objectives: [
      'Rewrite a vague question as a specific, measurable one.',
      'Tell a dimension (a way to slice) from a measure (a number to aggregate), and a metric from a KPI.',
      'Make decisions in a formula with IF, IFS, AND and OR.',
      'Count, total and average with conditions using COUNTIF(S), SUMIF(S), AVERAGEIF(S) and MAXIFS.',
      'Write criteria with operators, wildcards and cell references.',
    ],
    sections: [
      { heading: 'Vague questions do not have answers', html: `
        <p>"How is the school store doing?" cannot be answered, because "doing" is not defined. Analysts turn vague questions into <strong>measurable</strong> ones by naming exactly what to count, over what time, for which group, compared with what.</p>
        <table class="notes-table">
          <thead><tr><th>Vague</th><th>Measurable</th></tr></thead>
          <tbody>
            <tr><td>How is the store doing?</td><td>What was total revenue per week from January to April, and is it going up or down?</td></tr>
            <tr><td>Do snacks or drinks sell better?</td><td>What is total revenue from Snacks vs Drinks, and how many orders did each have?</td></tr>
            <tr><td>Who is our best cashier?</td><td>Which cashier processed the most orders, and what was their average order value?</td></tr>
          </tbody>
        </table>
        <div class="callout"><strong>Checklist for a measurable question:</strong> a number (count, total, average, percent), a subject (orders, students), a group or slice (by category, by month), a time frame, and ideally a comparison (vs last month, vs a goal).</div>` },
      { heading: 'Dimensions, measures, metrics and KPIs', html: `
        <ul>
          <li>A <strong>measure</strong> is a numeric field you aggregate: Qty, Price, Revenue, Hours.</li>
          <li>A <strong>dimension</strong> is a field you slice by: Category, Cashier, Payment, Month. Dimensions are usually text or dates.</li>
          <li>A <strong>metric</strong> is a calculated number that describes performance: total revenue, number of orders, average order value.</li>
          <li>A <strong>KPI</strong> (key performance indicator) is a metric the organization has decided matters most, usually with a <strong>target</strong>: "average order value of at least $8".</li>
        </ul>
        <p>Every summary you build from here on, in SUMIFS, pivot tables or Power BI, is the same idea: <em>aggregate a measure, sliced by a dimension</em>.</p>` },
      { heading: 'Helper columns', html: `
        <p>The store's data has Qty and Price but no Revenue. Rather than compute it inside every summary formula, add a <strong>helper column</strong>: in J2 type <code>=F2*G2</code> and fill it down. Now Revenue is a measure you can sum, average and chart. Name the header clearly (<code>Revenue</code>) so the next person understands it.</p>` },
      { heading: 'IF: a formula that decides', html: `
        <p><code>=IF(logical_test, value_if_true, value_if_false)</code></p>
        <ul>
          <li><code>=IF(J2&gt;=10,"Large","Small")</code> labels each order.</li>
          <li><code>=IF(H2="Card",J2*0.03,0)</code> computes a 3% card fee only for card payments.</li>
          <li>Combine tests with <code>AND(…)</code> (all must be true) and <code>OR(…)</code> (any can be true): <code>=IF(AND(E2="Apparel",F2&gt;1),"Bulk","")</code>.</li>
          <li>For several bands, <code>IFS</code> is easier to read than nested IFs: <code>=IFS(J2&gt;=20,"High",J2&gt;=8,"Medium",TRUE,"Low")</code>. Conditions are checked in order, and <code>TRUE</code> catches everything left.</li>
        </ul>
        <div class="callout tip">Text in a formula always goes in double quotes. Numbers never do. <code>=IF(F2&gt;"5",…)</code> compares with the <em>text</em> "5" and gives surprising results.</div>` },
      { heading: 'Counting, totaling and averaging with conditions', html: `
        <table class="notes-table">
          <thead><tr><th>Function</th><th>What it does</th><th>Example</th></tr></thead>
          <tbody>
            <tr><td><code>COUNTIF(range, criteria)</code></td><td>Count cells that match</td><td><code>=COUNTIF(E2:E61,"Snacks")</code></td></tr>
            <tr><td><code>SUMIF(range, criteria, sum_range)</code></td><td>Total the matching rows</td><td><code>=SUMIF(E2:E61,"Snacks",J2:J61)</code></td></tr>
            <tr><td><code>AVERAGEIF(range, criteria, avg_range)</code></td><td>Average the matching rows</td><td><code>=AVERAGEIF(H2:H61,"Card",J2:J61)</code></td></tr>
            <tr><td><code>COUNTIFS(r1, c1, r2, c2, …)</code></td><td>Count rows matching <em>all</em> conditions</td><td><code>=COUNTIFS(E2:E61,"Drinks",H2:H61,"Cash")</code></td></tr>
            <tr><td><code>SUMIFS(sum_range, r1, c1, r2, c2, …)</code></td><td>Total rows matching all conditions</td><td><code>=SUMIFS(J2:J61,E2:E61,"Drinks",I2:I61,"Sam")</code></td></tr>
            <tr><td><code>AVERAGEIFS</code>, <code>MAXIFS</code>, <code>MINIFS</code></td><td>Same pattern</td><td><code>=MAXIFS(J2:J61,E2:E61,"Apparel")</code></td></tr>
          </tbody>
        </table>
        <p>Watch the argument order: in <code>SUMIF</code> the sum range comes <em>last</em>; in <code>SUMIFS</code> it comes <em>first</em>.</p>` },
      { heading: 'Writing criteria', html: `
        <ul>
          <li>Exact match: <code>"Snacks"</code> (not case-sensitive).</li>
          <li>Comparisons go inside the quotes: <code>"&gt;=10"</code>, <code>"&lt;&gt;Cash"</code>.</li>
          <li>Wildcards: <code>*</code> any characters, <code>?</code> one character. <code>"*Bar"</code> matches Granola Bar.</li>
          <li>Use a cell instead of typing the value: <code>=SUMIFS(J$2:J$61,E$2:E$61,L2)</code>. Now one formula fills down a whole summary table.</li>
          <li>Combine an operator with a cell: <code>"&gt;="&amp;M1</code>. Dates work the same way: <code>"&gt;="&amp;DATE(2026,3,1)</code>.</li>
        </ul>` },
      { heading: 'Build a summary table', html: `
        <p>The pattern you will use all semester: list the dimension values down a column (Snacks, Drinks, Supplies, Apparel), then write one SUMIFS that refers to the label cell and fill it down. Lock the data ranges with <code>$</code> so they do not move. In week 10 you will build the same table with a pivot in seconds, and you will understand exactly what the pivot is doing.</p>` },
    ],
    keyTerms: [
      ['Measurable question', 'A question that names a number, subject, group and time frame.'],
      ['Measure', 'A numeric field that gets aggregated.'],
      ['Dimension', 'A field used to group or slice data.'],
      ['KPI', 'A key metric with a target, chosen because it matters to a goal.'],
      ['Helper column', 'A calculated column added to make later analysis simpler.'],
      ['Criteria', 'The condition a COUNTIF/SUMIF-style function matches against.'],
    ],
    practice: [
      {
        id: 'p6a', title: 'Workbook 6A · Store KPIs with conditional functions',
        intro: 'First add the Revenue helper column: J2 = F2*G2, filled down to J61. Then answer each question in column M with one formula. Lock ranges with $ where it helps.',
        tools: 'basic',
        workbook: {
          sheets: [(() => {
            const s = D.storeSalesSheet('Sales', { extraHeaders: ['Revenue'], formats: { ['J2:J' + N]: 'currency', M2: 'currency', M3: 'currency', M5: 'currency', M6: 'currency', M7: 'currency' }, widths: { J: 80, L: 330, M: 90 } });
            s.cells = {
              L1: { input: 'Question', style: { bold: true } }, M1: { input: 'Answer', style: { bold: true } },
              L2: 'Total revenue', L3: 'Revenue from Snacks', L4: 'Number of Drinks orders', L5: 'Revenue from Drinks paid with Cash',
              L6: 'Average revenue per order', L7: 'Largest single Apparel order (revenue)', L8: 'Orders with Qty of 4 or more',
            };
            s.lock = ['A1:I' + N, 'L1:M1', 'L2:L8'];
            s.styles = { ['J2:J' + N]: { fill: 'answer' }, 'M2:M8': { fill: 'answer' } };
            return s;
          })()],
        },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Sales', range: 'J2:J' + N, solution: '=F2*G2', requireFormula: true, points: 3, prompt: 'Revenue helper column J2:J61 (= Qty × Price).' },
          { id: 'q2', type: 'cell', sheet: 'Sales', cell: 'M2', solution: '=SUM(J2:J61)', requireFormula: true, prompt: 'M2: total revenue.' },
          { id: 'q3', type: 'cell', sheet: 'Sales', cell: 'M3', solution: '=SUMIF(E2:E61,"Snacks",J2:J61)', functions: [['SUMIF', 'SUMIFS']], prompt: 'M3: Snacks revenue with SUMIF.' },
          { id: 'q4', type: 'cell', sheet: 'Sales', cell: 'M4', solution: '=COUNTIF(E2:E61,"Drinks")', functions: [['COUNTIF', 'COUNTIFS']], prompt: 'M4: number of Drinks orders with COUNTIF.' },
          { id: 'q5', type: 'cell', sheet: 'Sales', cell: 'M5', solution: '=SUMIFS(J2:J61,E2:E61,"Drinks",H2:H61,"Cash")', functions: ['SUMIFS'], prompt: 'M5: Drinks revenue paid in Cash, with SUMIFS (two conditions).' },
          { id: 'q6', type: 'cell', sheet: 'Sales', cell: 'M6', solution: '=AVERAGE(J2:J61)', requireFormula: true, tolerance: 0.005, prompt: 'M6: average revenue per order.' },
          { id: 'q7', type: 'cell', sheet: 'Sales', cell: 'M7', solution: '=MAXIFS(J2:J61,E2:E61,"Apparel")', functions: ['MAXIFS'], prompt: 'M7: largest Apparel order with MAXIFS.' },
          { id: 'q8', type: 'cell', sheet: 'Sales', cell: 'M8', solution: '=COUNTIF(F2:F61,">=4")', functions: [['COUNTIF', 'COUNTIFS']], prompt: 'M8: orders with Qty ≥ 4 (criteria ">=4").' },
        ],
      },
      {
        id: 'p6b', title: 'Workbook 6B · IF, IFS and AND',
        intro: 'Label each order. Write the formula in row 2, then Fill ↓ to row 61.',
        tools: 'basic',
        workbook: {
          sheets: [(() => {
            const s = salesWithRevenue('Sales');
            s.rows[0] = [...s.rows[0], 'Size', 'Band', 'Bulk Apparel?'];
            s.widths = { ...s.widths, K: 70, L: 80, M: 110 };
            s.styles = { ['K2:M' + N]: { fill: 'answer' } };
            s.lock = ['A1:J' + N, 'K1:M1'];
            return s;
          })()],
        },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Sales', range: 'K2:K' + N, solution: '=IF(J2>=10,"Large","Small")', functions: ['IF'], points: 2, prompt: 'K: "Large" when Revenue is at least $10, otherwise "Small".' },
          { id: 'q2', type: 'range', sheet: 'Sales', range: 'L2:L' + N, solution: '=IFS(J2>=20,"High",J2>=8,"Medium",TRUE,"Low")', functions: [['IFS', 'IF']], points: 2, prompt: 'L: "High" (≥ $20), "Medium" (≥ $8) or "Low", using IFS.' },
          { id: 'q3', type: 'range', sheet: 'Sales', range: 'M2:M' + N, solution: '=IF(AND(E2="Apparel",F2>1),"Yes","No")', functions: ['AND'], points: 2, prompt: 'M: "Yes" when Category is Apparel AND Qty is more than 1, else "No".' },
        ],
      },
    ],
    assignment: {
      id: 'hw-6', title: 'KPI sheet for the school store', points: 25,
      html: `<p>The store manager asked, "How is the store doing?" Turn that into measurable KPIs.</p>
        <ol><li>On the <code>KPIs</code> sheet, rewrite the manager's question as one measurable question in <code>B2</code>.</li>
        <li>Fill the category summary B6:D9 with <strong>one</strong> formula per column that refers to the label in column A, filled down. Use <code>$</code> to lock the data ranges.</li>
        <li>Fill the KPI cells B12:B14.</li>
        <li>In <code>B16</code>, write one sentence of advice for the manager based on your numbers.</li></ol>`,
      tools: 'basic',
      workbook: {
        sheets: [
          salesWithRevenue('Sales'),
          {
            name: 'KPIs',
            rows: [
              ['School store KPIs', null, null, null],
              ['Measurable question', null, null, null],
              [null, null, null, null],
              [null, null, null, null],
              ['Category', 'Orders', 'Revenue', 'Avg revenue per order'],
              ['Snacks'], ['Drinks'], ['Supplies'], ['Apparel'],
              [null], [null],
              ['Total revenue'], ['Share of revenue from Apparel'], ['Card orders as % of all orders'],
              [null], ['Advice'],
            ],
            formats: { 'C6:D9': 'currency', B12: 'currency', B13: 'percent1', B14: 'percent1' },
            lock: ['A1:A16', 'B5:D5'],
            widths: { A: 230, B: 110, C: 110, D: 160 },
            styles: { A1: { bold: true }, B2: { fill: 'answer' }, 'B6:D9': { fill: 'answer' }, 'B12:B14': { fill: 'answer' }, B16: { fill: 'answer' } },
          },
        ],
      },
      questions: [
        { id: 'a1', type: 'filled', sheet: 'KPIs', range: 'B2', min: 1, points: 3, prompt: 'Measurable question in B2.' },
        { id: 'a2', type: 'range', sheet: 'KPIs', range: 'B6:B9', solution: '=COUNTIF(Sales!$E$2:$E$61,A6)', requireFormula: true, points: 4, prompt: 'Orders per category (B6:B9).' },
        { id: 'a3', type: 'range', sheet: 'KPIs', range: 'C6:C9', solution: '=SUMIF(Sales!$E$2:$E$61,A6,Sales!$J$2:$J$61)', requireFormula: true, points: 4, prompt: 'Revenue per category (C6:C9).' },
        { id: 'a4', type: 'range', sheet: 'KPIs', range: 'D6:D9', solution: '=C6/B6', requireFormula: true, tolerance: 0.005, points: 4, prompt: 'Average revenue per order per category (D6:D9).' },
        { id: 'a5', type: 'cell', sheet: 'KPIs', cell: 'B12', solution: '=SUM(Sales!J2:J61)', requireFormula: true, points: 2, prompt: 'Total revenue.' },
        { id: 'a6', type: 'cell', sheet: 'KPIs', cell: 'B13', solution: '=C9/B12', requireFormula: true, tolerance: 0.0005, points: 3, prompt: 'Apparel share of revenue.' },
        { id: 'a7', type: 'cell', sheet: 'KPIs', cell: 'B14', solution: '=COUNTIF(Sales!H2:H61,"Card")/COUNTA(Sales!A2:A61)', requireFormula: true, tolerance: 0.0005, points: 3, prompt: 'Card orders as a percent of all orders.' },
        { id: 'a8', type: 'filled', sheet: 'KPIs', range: 'B16', min: 1, points: 2, prompt: 'One sentence of advice in B16.' },
      ],
      rubric: [['The question in B2 names a number, a group and a time frame', 'Teacher check'], ['The advice follows from the numbers', 'Teacher check']],
    },
    quiz: null,
  });

  // ---------------------------------------------------------------- Week 7
  const studentsSheet = { name: 'Students', rows: D.students, lock: true, widths: { E: 90 } };
  const coursesSheet = { name: 'Courses', rows: D.courses, lock: true, widths: { B: 130, C: 90 } };

  C.lessons.push({
    week: 7, unit: 2, id: 'week-7',
    title: 'Lookups and relational thinking',
    summary: 'Why data is split across tables, primary and foreign keys, one-to-many relationships and referential integrity, and joining tables in Excel with XLOOKUP, VLOOKUP and INDEX/MATCH.',
    objectives: [
      'Explain why organizations store data in several related tables instead of one big one.',
      'Identify the primary key of a table and test that it is unique.',
      'Identify a foreign key and describe the one-to-many relationship it creates.',
      'Find orphan records that break referential integrity.',
      'Join tables with XLOOKUP (and read VLOOKUP and INDEX/MATCH), handling missing matches.',
    ],
    sections: [
      { heading: 'Why split data into tables?', html: `
        <p>Imagine storing every enrollment as one row with the student's name, grade, advisor, course name, teacher and credits. Maya's advisor would be typed once for <em>every</em> course she takes. If her advisor changes, you must find and fix every copy, and the copies you miss make the data inconsistent. This is called <strong>redundancy</strong>, and it causes <strong>update anomalies</strong>.</p>
        <p>Databases avoid this by storing each kind of thing in its own table, once: a Students table, a Courses table, and an Enrollments table that connects them. This idea is called <strong>normalization</strong>, and a database built this way is a <strong>relational database</strong>.</p>` },
      { heading: 'Primary keys', html: `
        <p>A <strong>primary key (PK)</strong> is the column (or set of columns) that uniquely identifies each row in a table. Rules:</p>
        <ul>
          <li><strong>Unique:</strong> no two rows share it.</li>
          <li><strong>Never blank.</strong></li>
          <li><strong>Stable:</strong> it should not change. A student ID is a good key; a name is a bad one (two students can share a name, and names change).</li>
        </ul>
        <p>Examples: <code>StudentID</code> in Students, <code>CourseID</code> in Courses, <code>EnrollmentID</code> in Enrollments. When no single column is unique, two columns together can form a <strong>composite key</strong>, e.g. StudentID + CourseID.</p>
        <div class="callout tip"><strong>Test a key in Excel:</strong> next to each ID, <code>=COUNTIF($A$2:$A$11,A2)</code>. Every result must be 1. Any 2 or more means the column is not a valid primary key.</div>` },
      { heading: 'Foreign keys and relationships', html: `
        <p>A <strong>foreign key (FK)</strong> is a column in one table that holds the primary key of another table. In Enrollments, <code>StudentID</code> is a foreign key pointing to Students, and <code>CourseID</code> is a foreign key pointing to Courses.</p>
        <p>This creates a <strong>one-to-many relationship</strong>: one student can have many enrollments, but each enrollment belongs to exactly one student. Diagrams draw it as <code>Students 1 ──── ∞ Enrollments</code>.</p>
        <table class="notes-table">
          <thead><tr><th>Table</th><th>Primary key</th><th>Foreign keys</th></tr></thead>
          <tbody>
            <tr><td>Students</td><td>StudentID</td><td>none</td></tr>
            <tr><td>Courses</td><td>CourseID</td><td>none</td></tr>
            <tr><td>Enrollments</td><td>EnrollmentID</td><td>StudentID → Students, CourseID → Courses</td></tr>
          </tbody>
        </table>
        <p><strong>Referential integrity</strong> means every foreign key value actually exists in the table it points to. An enrollment for student <code>S1012</code> when no S1012 exists in Students is an <strong>orphan record</strong>: whose score is it? Databases can block orphans; in a spreadsheet, a lookup that returns <code>#N/A</code> is how you find them.</p>
        <p>You will see these same relationships again in week 12 (Power Query merges) and week 13 (the Power BI model view, where you literally draw the 1-to-many lines).</p>` },
      { heading: 'XLOOKUP: joining tables in Excel', html: `
        <p><code>=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])</code></p>
        <p>"Find this key in that column and bring back the value from this other column." To add each student's last name to the Enrollments sheet:</p>
        <pre class="raw">=XLOOKUP(B2, Students!$A$2:$A$11, Students!$C$2:$C$11, "Not found")</pre>
        <ul>
          <li><code>B2</code> is the foreign key in this row.</li>
          <li><code>Students!$A$2:$A$11</code> is the primary key column in the other table.</li>
          <li><code>Students!$C$2:$C$11</code> is what to return. It must be the same size as the lookup array.</li>
          <li>The fourth argument replaces <code>#N/A</code> when there is no match. Leave it out while you are checking data quality, because the #N/A errors <em>are</em> the finding.</li>
        </ul>` },
      { heading: 'VLOOKUP and INDEX/MATCH', html: `
        <p>You will meet older workbooks that use these, so you need to read them.</p>
        <p><code>=VLOOKUP(B2, Students!$A$2:$E$11, 3, FALSE)</code> looks for B2 in the <em>first</em> column of the table and returns the 3rd column. Always use <code>FALSE</code> (exact match) for IDs; the default approximate match silently returns wrong rows. VLOOKUP cannot look to the left, and breaks if someone inserts a column.</p>
        <p><code>=INDEX(Students!$C$2:$C$11, MATCH(B2, Students!$A$2:$A$11, 0))</code>: MATCH finds the position of the key, INDEX returns the value at that position. It works in every Excel version and in any direction.</p>` },
      { heading: 'Common lookup problems', html: `
        <ul>
          <li><strong>#N/A for a key you can see:</strong> extra spaces (<code>"S1003 "</code>), or a number stored as text in one table and a number in the other. Fix with TRIM or by converting the type (week 8).</li>
          <li><strong>Wrong value returned:</strong> VLOOKUP without FALSE, or a return range that does not line up with the lookup range.</li>
          <li><strong>Duplicates in the lookup column:</strong> lookups return only the first match, which is why the lookup side must be a primary key.</li>
        </ul>` },
    ],
    keyTerms: [
      ['Primary key (PK)', 'A column whose value uniquely identifies each row.'],
      ['Foreign key (FK)', 'A column that holds another table\'s primary key, linking the tables.'],
      ['Composite key', 'Two or more columns that together uniquely identify a row.'],
      ['One-to-many', 'One row in a table relates to many rows in another.'],
      ['Referential integrity', 'Every foreign key value exists as a primary key in the related table.'],
      ['Orphan record', 'A row whose foreign key points to nothing.'],
      ['Normalization', 'Organizing data into related tables so each fact is stored once.'],
      ['XLOOKUP', 'Finds a key in one column and returns the matching value from another.'],
    ],
    practice: [
      {
        id: 'p7a', title: 'Workbook 7A · Keys and lookups',
        intro: 'Enrollments has foreign keys to Students and Courses. Use XLOOKUP to bring in names, then find the orphans. Leave out the "if not found" argument in E and F so missing keys show #N/A.',
        tools: 'basic',
        workbook: {
          sheets: [
            {
              name: 'Enrollments',
              rows: [[...D.grades[0], 'Last Name', 'Course Name', 'Orphan?'], ...D.grades.slice(1)],
              lock: 'A1:G' + D.grades.length,
              widths: { A: 100, E: 100, F: 130, G: 80, I: 300, J: 80 },
              cells: { I1: { input: 'Key check', style: { bold: true }, locked: true }, I2: { input: 'Times S1003 appears in Enrollments', locked: true }, I3: { input: 'Is StudentID unique in Students? (max COUNTIF)', locked: true }, I4: { input: 'Number of orphan enrollments', locked: true } },
              styles: { ['E2:G' + D.grades.length]: { fill: 'answer' }, 'J2:J4': { fill: 'answer' } },
            },
            studentsSheet, coursesSheet,
          ],
        },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Enrollments', range: 'E2:E' + D.grades.length, solution: '=XLOOKUP(B2,Students!$A$2:$A$11,Students!$C$2:$C$11)', functions: [['XLOOKUP', 'VLOOKUP', 'INDEX']], points: 3, prompt: 'E: each student\'s last name via XLOOKUP on StudentID (orphans show #N/A).' },
          { id: 'q2', type: 'range', sheet: 'Enrollments', range: 'F2:F' + D.grades.length, solution: '=XLOOKUP(C2,Courses!$A$2:$A$5,Courses!$B$2:$B$5)', functions: [['XLOOKUP', 'VLOOKUP', 'INDEX']], points: 3, prompt: 'F: course name via XLOOKUP on CourseID.' },
          { id: 'q3', type: 'range', sheet: 'Enrollments', range: 'G2:G' + D.grades.length, solution: '=IF(ISNA(E2),"Yes","No")', requireFormula: true, points: 2, prompt: 'G: "Yes" when the student lookup failed, otherwise "No" (hint: ISNA).' },
          { id: 'q4', type: 'cell', sheet: 'Enrollments', cell: 'J2', solution: '=COUNTIF(B2:B17,"S1003")', requireFormula: true, prompt: 'J2: how many enrollments S1003 has (one student, many enrollments).' },
          { id: 'q5', type: 'cell', sheet: 'Enrollments', cell: 'J3', expect: 1, solution: 1, prompt: 'J3: the largest COUNTIF of any StudentID within Students!A2:A11 (1 means unique).' },
          { id: 'q6', type: 'cell', sheet: 'Enrollments', cell: 'J4', solution: '=COUNTIF(G2:G17,"Yes")', requireFormula: true, prompt: 'J4: number of orphan enrollments.' },
          { id: 'q7', type: 'choice', prompt: 'In the Enrollments table, StudentID is…', options: ['the primary key of Enrollments', 'a foreign key that points to Students', 'a composite key', 'not a key at all'], answer: 1 },
          { id: 'q8', type: 'choice', prompt: 'Which relationship is this: one student, many enrollments?', options: ['One-to-one', 'One-to-many', 'Many-to-many', 'No relationship'], answer: 1 },
        ],
      },
    ],
    assignment: {
      id: 'hw-7', title: 'Combine the order tables', points: 25,
      html: `<p>The store's online order form saves only a ProductID and a quantity. Combine it with the product catalog, and report any orders whose ProductID breaks referential integrity.</p>
        <ol><li>On <code>Orders</code>, fill Item (C), Price (D) and Revenue (E) for every row. For Item, show <code>Not found</code> when the ProductID is not in the catalog. For Price, return <code>0</code> when not found.</li>
        <li>In H2 and H3, count the bad orders and total the revenue.</li>
        <li>On <code>Keys</code>, answer the three questions.</li></ol>`,
      tools: 'basic',
      workbook: {
        sheets: [
          {
            name: 'Orders',
            rows: [
              ['OrderID', 'ProductID', 'Item', 'Price', 'Revenue', null, null, null],
              ['W01', 'P202', null, null, null, null, 'Orders with an unknown ProductID'], ['W02', 'P101', null, null, null, null, 'Total revenue'],
              ['W03', 'P401'], ['W04', 'P1O2'], ['W05', 'P303'], ['W06', 'P202'], ['W07', 'P999'], ['W08', 'P402'], ['W09', 'P301'], ['W10', 'P203'], ['W11', 'p101'], ['W12', 'P102'],
            ].map((r) => [r[0], r[1], null, null, null, null, r[6] || null]),
            cells: { F1: { input: 'Qty', style: { bold: true } }, ...Object.fromEntries([2, 1, 1, 3, 2, 4, 1, 1, 5, 6, 2, 1].map((q, i) => ['F' + (i + 2), { input: q, locked: true }])) },
            formats: { 'D2:E13': 'currency', H3: 'currency' },
            lock: ['A1:F13', 'G2:G3'],
            widths: { C: 120, G: 230 },
            styles: { 'C2:E13': { fill: 'answer' }, 'H2:H3': { fill: 'answer' } },
          },
          { name: 'Products', rows: [['ProductID', 'Item', 'Category', 'Price'], ...D.catalog], formats: { 'D2:D13': 'currency' }, lock: true, widths: { B: 120 } },
          {
            name: 'Keys',
            rows: [['Question', 'Answer'], ['What is the primary key of Products?'], ['Which column in Orders is a foreign key, and which table does it point to?'], ['Explain why W04 and W11 did not match, and how you would prevent it at data entry (hint: week 5).']],
            lock: 'A1:A4', widths: { A: 460, B: 420 }, styles: { 'B2:B4': { fill: 'answer' } },
          },
        ],
      },
      questions: [
        { id: 'a1', type: 'range', sheet: 'Orders', range: 'C2:C13', solution: '=XLOOKUP(B2,Products!$A$2:$A$13,Products!$B$2:$B$13,"Not found")', functions: [['XLOOKUP', 'VLOOKUP', 'INDEX', 'IFERROR', 'IFNA']], points: 5, prompt: 'Item looked up, "Not found" for bad keys.', hint: 'Use the fourth argument of XLOOKUP, or wrap VLOOKUP in IFNA.' },
        { id: 'a2', type: 'range', sheet: 'Orders', range: 'D2:D13', solution: '=XLOOKUP(B2,Products!$A$2:$A$13,Products!$D$2:$D$13,0)', requireFormula: true, points: 5, prompt: 'Price looked up, 0 for bad keys.' },
        { id: 'a3', type: 'range', sheet: 'Orders', range: 'E2:E13', solution: '=D2*F2', requireFormula: true, points: 3, prompt: 'Revenue = Price × Qty.' },
        { id: 'a4', type: 'cell', sheet: 'Orders', cell: 'H2', solution: '=COUNTIF(C2:C13,"Not found")', requireFormula: true, points: 3, prompt: 'Count of orders with an unknown ProductID.' },
        { id: 'a5', type: 'cell', sheet: 'Orders', cell: 'H3', solution: '=SUM(E2:E13)', requireFormula: true, points: 3, prompt: 'Total revenue.' },
        { id: 'a6', type: 'filled', sheet: 'Keys', range: 'B2:B4', min: 3, points: 6, prompt: 'All three key questions answered.' },
      ],
      rubric: [['Keys: ProductID is the PK of Products; Orders.ProductID is an FK to Products', 'Teacher check'], ['W04 uses the letter O instead of zero; W11 is lower case (XLOOKUP is not case-sensitive, so it matches, but a database would not). A list validation prevents both.', 'Teacher check']],
    },
    quiz: 'quiz-7',
  });

  // ---------------------------------------------------------------- Week 8
  const messy = [
    ['Raw Name', 'Raw Email', 'Signup Date', 'Phone'],
    ['  maya PATEL', 'MPatel@School.org ', '2/3/2026', '(555) 014-1000'],
    ['luis garcia ', 'lgarcia@school.org', '2/3/2026', '555.017.7000'],
    ['AISHA  BELLO', ' abello@SCHOOL.org', '2/4/2026', '555-011-2000'],
    ['ethan nguyen', 'enguyen@school.org', '2/9/2026', '(555) 019-0000'],
    ['Grace Kim', 'GKIM@school.org', '2/10/2026', '555 016 4000'],
    ['noah    smith', 'nsmith@school.org  ', '2/12/2026', '555-013-3000'],
    ['SOFIA ROSSI', 'srossi@school.org', '2/16/2026', '(555) 018-8000'],
    ['  kai tanaka  ', 'ktanaka@school.org', '2/20/2026', '555.015.5000'],
  ];

  C.lessons.push({
    week: 8, unit: 2, id: 'week-8',
    title: 'Cleaning text and dates',
    summary: 'Standardizing messy text with TRIM, PROPER, UPPER and LOWER; splitting and joining with LEFT, RIGHT, MID, FIND, LEN, SUBSTITUTE and TEXTJOIN; working with dates; removing duplicates; and why doing all this by hand every month does not scale.',
    objectives: [
      'Standardize text with TRIM, PROPER, UPPER, LOWER and SUBSTITUTE.',
      'Extract parts of text with LEFT, RIGHT, MID, FIND and LEN, and join text with & and TEXTJOIN.',
      'Extract date parts with YEAR, MONTH, DAY and WEEKDAY, build dates with DATE, and format them with TEXT.',
      'Convert formula results to values, and remove duplicate rows.',
      'Explain why a repeatable cleaning process beats cleaning by hand.',
    ],
    sections: [
      { heading: 'Messy data is normal', html: `
        <p>Data typed by people is inconsistent: extra spaces, random capitals, different phone formats, dates stored as text. Analysts often say 60–80% of their time goes to cleaning. The good news: almost every problem has a function.</p>
        <p><strong>The golden rule:</strong> never overwrite the raw data. Add clean columns next to it with formulas, check them, then (if needed) copy and <strong>Paste Values</strong> over a new sheet. Keep the original so you can always start over.</p>` },
      { heading: 'Standardizing text', html: `
        <table class="notes-table">
          <thead><tr><th>Function</th><th>Result</th></tr></thead>
          <tbody>
            <tr><td><code>=TRIM("  maya   PATEL ")</code></td><td><code>maya PATEL</code>: removes leading/trailing spaces and squeezes inside spaces to one</td></tr>
            <tr><td><code>=PROPER("maya PATEL")</code></td><td><code>Maya Patel</code></td></tr>
            <tr><td><code>=UPPER("p101")</code> / <code>=LOWER("MPatel@School.org")</code></td><td><code>P101</code> / <code>mpatel@school.org</code></td></tr>
            <tr><td><code>=SUBSTITUTE("555.017.7000",".","-")</code></td><td><code>555-017-7000</code></td></tr>
            <tr><td><code>=CLEAN(A2)</code></td><td>Removes invisible non-printing characters from exports</td></tr>
          </tbody>
        </table>
        <p>Functions nest: <code>=PROPER(TRIM(A2))</code> trims first, then capitalizes. Read nested formulas from the inside out.</p>` },
      { heading: 'Splitting and joining text', html: `
        <ul>
          <li><code>LEFT(text, n)</code>, <code>RIGHT(text, n)</code>, <code>MID(text, start, n)</code> take characters from the left, right or middle.</li>
          <li><code>LEN(text)</code> counts characters; <code>FIND(find, text)</code> returns the position of the first match.</li>
          <li>First name: <code>=LEFT(B2, FIND(" ", B2) - 1)</code>. Last name: <code>=MID(B2, FIND(" ", B2) + 1, 100)</code>.</li>
          <li>Email username: <code>=LEFT(C2, FIND("@", C2) - 1)</code>.</li>
          <li>Join with <code>&amp;</code>: <code>=E2 &amp; ", " &amp; D2</code> gives "Patel, Maya". <code>TEXTJOIN(", ", TRUE, A2:C2)</code> joins a range and skips blanks.</li>
        </ul>
        <div class="callout">When a pattern is simple (always one comma), Text to Columns from week 4 is faster. Formulas win when the data will change, because they recalculate.</div>` },
      { heading: 'Working with dates', html: `
        <p>Because dates are numbers (week 2), you can do math with them and pull out parts:</p>
        <table class="notes-table">
          <thead><tr><th>Formula</th><th>For 3/15/2026 (a Sunday)</th></tr></thead>
          <tbody>
            <tr><td><code>=YEAR(A2)</code>, <code>=MONTH(A2)</code>, <code>=DAY(A2)</code></td><td>2026, 3, 15</td></tr>
            <tr><td><code>=WEEKDAY(A2)</code></td><td>1 (Sunday = 1 … Saturday = 7)</td></tr>
            <tr><td><code>=TEXT(A2,"mmm")</code>, <code>=TEXT(A2,"dddd")</code></td><td>Mar, Sunday</td></tr>
            <tr><td><code>=DATE(2026,3,15)</code></td><td>Builds a real date from parts</td></tr>
            <tr><td><code>=A2+30</code>, <code>=B2-A2</code></td><td>30 days later; days between two dates</td></tr>
            <tr><td><code>=DATEDIF(A2,TODAY(),"y")</code></td><td>Whole years between two dates (age)</td></tr>
          </tbody>
        </table>
        <p>A date stored as <em>text</em> ("2026-03-15" left-aligned) will not work with these. <code>=DATEVALUE</code> or <code>=VALUE</code> converts many of them, or rebuild it with <code>DATE(LEFT(…),MID(…),RIGHT(…))</code>.</p>` },
      { heading: 'Duplicates', html: `
        <p>Find duplicates before removing them: <code>=COUNTIF($A$2:$A$100,A2)&gt;1</code> flags repeated values, and conditional formatting can highlight them. Then use <strong>Data › Remove Duplicates</strong>, choosing which columns define a duplicate. Removing duplicates by the wrong columns deletes real data: two different students can share a name.</p>` },
      { heading: 'Now do it again next month', html: `
        <p>Here is the problem. You spent a class period cleaning February's sign-ups. In March, a new file arrives with the same problems, and you do every step again. And again in April. Every repetition is a chance for a mistake, and nobody can see what you did.</p>
        <p>In week 11 you will record these steps once in <strong>Power Query</strong>, and next month you will press <strong>Refresh</strong>. Keep a list of the cleaning steps you do this week; you will rebuild them there.</p>` },
    ],
    keyTerms: [
      ['TRIM', 'Removes extra spaces.'],
      ['PROPER / UPPER / LOWER', 'Change capitalization.'],
      ['SUBSTITUTE', 'Replaces specific text with other text.'],
      ['LEFT / RIGHT / MID', 'Extract characters from text.'],
      ['Concatenate', 'Join text together (&, CONCAT, TEXTJOIN).'],
      ['Paste Values', 'Replace formulas with their results.'],
    ],
    practice: [
      {
        id: 'p8a', title: 'Workbook 8A · Clean the sign-up list',
        intro: 'Raw data stays in A:D. Write each clean column in row 2, then Fill ↓ to row 9.',
        tools: 'data',
        workbook: {
          sheets: [{
            name: 'Signups',
            rows: messy.map((r, i) => (i === 0 ? [...r, 'Clean Name', 'First', 'Last', 'Clean Email', 'Username', 'Clean Phone', 'Signup Day'] : r)),
            lock: ['A1:K1', 'A2:D9'],
            formats: { 'C2:C9': 'date' },
            widths: { A: 120, B: 160, C: 90, D: 110, E: 110, F: 70, G: 70, H: 160, I: 80, J: 105, K: 85 },
            styles: { 'E2:K9': { fill: 'answer' } },
          }],
        },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Signups', range: 'E2:E9', solution: '=PROPER(TRIM(A2))', functions: ['TRIM', 'PROPER'], points: 2, prompt: 'E: name trimmed and in Proper Case.' },
          { id: 'q2', type: 'range', sheet: 'Signups', range: 'F2:F9', solution: '=LEFT(E2,FIND(" ",E2)-1)', requireFormula: true, points: 2, prompt: 'F: first name from E.' },
          { id: 'q3', type: 'range', sheet: 'Signups', range: 'G2:G9', solution: '=MID(E2,FIND(" ",E2)+1,100)', requireFormula: true, points: 2, prompt: 'G: last name from E.' },
          { id: 'q4', type: 'range', sheet: 'Signups', range: 'H2:H9', solution: '=LOWER(TRIM(B2))', functions: ['LOWER', 'TRIM'], points: 2, prompt: 'H: email trimmed and lower case.' },
          { id: 'q5', type: 'range', sheet: 'Signups', range: 'I2:I9', solution: '=LEFT(H2,FIND("@",H2)-1)', requireFormula: true, points: 2, prompt: 'I: username (the part before @).' },
          { id: 'q6', type: 'range', sheet: 'Signups', range: 'J2:J9', solution: '=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(D2,"(",""),")",""),".","-")," ","-")', requireFormula: true, points: 2, prompt: 'J: phone as 555-014-1000 (replace ( ) . and spaces).', hint: 'Nest SUBSTITUTE: remove "(" and ")", then turn "." and " " into "-". Watch the space after ")".' },
          { id: 'q7', type: 'range', sheet: 'Signups', range: 'K2:K9', solution: '=TEXT(C2,"dddd")', requireFormula: true, points: 2, prompt: 'K: weekday name of the signup date (TEXT with "dddd").' },
        ],
      },
    ],
    assignment: {
      id: 'hw-8', title: 'Clean the March sign-up list', points: 25,
      html: `<p>The March file arrived with the same problems plus duplicates. Clean it, and keep a written list of your steps; you will automate them in week 11.</p>
        <ol><li>Fill the clean columns D:F for every row (D = Proper-case trimmed name, E = lower-case trimmed email, F = month name of the date, e.g. "March").</li>
        <li>Now remove duplicates from A1:F13 <strong>using the clean columns</strong> (Name, Email) plus Signup Date. The raw rows differ by stray spaces, so Excel would not see them as duplicates; the clean columns do. That is why you clean first.</li>
        <li>In <code>I2</code>, count the clean sign-ups with a formula.</li>
        <li>On the <code>Steps</code> sheet, list each cleaning step you did, one per row (at least 4).</li></ol>`,
      tools: 'data',
      workbook: {
        sheets: [
          {
            name: 'March',
            rows: [
              ['Raw Name', 'Raw Email', 'Signup Date', 'Name', 'Email', 'Month', null, null, null],
              ['lily CHEN ', 'LChen@school.org', '3/2/2026', null, null, null, null, 'Clean sign-ups'],
              ['ben walker', 'bwalker@school.org ', '3/2/2026'],
              ['Lily Chen', 'lchen@school.org', '3/2/2026'],
              ['  CHLOE davis', 'cdavis@SCHOOL.org', '3/4/2026'],
              ['diego lopez', 'dlopez@school.org', '3/5/2026'],
              ['Ben Walker', 'BWalker@school.org', '3/2/2026'],
              ['hana sato', 'hsato@school.org', '3/9/2026'],
              ['IVY BROOKS', 'ibrooks@school.org', '3/11/2026'],
              ['jack   ellis', 'jellis@school.org', '3/12/2026'],
              ['hana sato', 'hsato@school.org', '3/9/2026'],
              ['Kira Novak', 'knovak@school.org', '3/16/2026'],
              ['mia rivera', 'mrivera@school.org', '3/18/2026'],
            ].map((r) => [r[0], r[1], r[2], null, null, null, null, r[7] || null]),
            cells: { D1: { input: 'Name', style: { bold: true } }, E1: { input: 'Email', style: { bold: true } }, F1: { input: 'Month', style: { bold: true } } },
            formats: { 'C2:C13': 'date' },
            widths: { A: 120, B: 160, D: 120, E: 160, H: 120 },
            lock: ['D1:F1', 'H2'],
            styles: { 'D2:F13': { fill: 'answer' }, I2: { fill: 'answer' } },
          },
          { name: 'Steps', rows: [['Step #', 'What I did and why'], [1], [2], [3], [4], [5], [6]], lock: 'A1:B1', widths: { B: 520 }, styles: { 'B2:B7': { fill: 'answer' } } },
        ],
      },
      model: (wb) => {
        for (let r = 2; r <= 13; r++) {
          wb.setInput('March', 'D' + r, '=PROPER(TRIM(A' + r + '))');
          wb.setInput('March', 'E' + r, '=LOWER(TRIM(B' + r + '))');
          wb.setInput('March', 'F' + r, '=TEXT(C' + r + ',"mmmm")');
        }
        wb.recalc();
        wb.removeDuplicates('March', 'A1:F13', true, [2, 3, 4]);
      },
      questions: [
        { id: 'a1', type: 'range', sheet: 'March', range: 'A4:A10', expect: [['  CHLOE davis'], ['diego lopez'], ['hana sato'], ['IVY BROOKS'], ['jack   ellis'], ['Kira Novak'], ['mia rivera']], exactText: true, points: 5, prompt: 'Duplicates removed: 9 unique sign-ups remain in A2:A10, in their original order (checked in A4:A10).' },
        { id: 'a2', type: 'range', sheet: 'March', range: 'D2:D10', solution: '=PROPER(TRIM(A2))', requireFormula: true, points: 4, prompt: 'Clean names.' },
        { id: 'a3', type: 'range', sheet: 'March', range: 'E2:E10', solution: '=LOWER(TRIM(B2))', requireFormula: true, points: 4, prompt: 'Clean emails.' },
        { id: 'a4', type: 'range', sheet: 'March', range: 'F2:F10', solution: '=TEXT(C2,"mmmm")', requireFormula: true, points: 3, prompt: 'Month name.' },
        { id: 'a5', type: 'cell', sheet: 'March', cell: 'I2', solution: '=COUNTA(A2:A13)', expect: 9, requireFormula: true, points: 3, prompt: 'Count of clean sign-ups.' },
        { id: 'a6', type: 'filled', sheet: 'Steps', range: 'B2:B7', min: 4, points: 6, prompt: 'At least four cleaning steps listed.' },
      ],
    },
    quiz: 'quiz-8',
  });

  // ---------------------------------------------------------------- Week 9
  const monthly = [['Month', 'Revenue'], ['Jan', 161.75], ['Feb', 114.25], ['Mar', 100.25], ['Apr', 195]];
  const byCategory = [['Category', 'Revenue'], ['Snacks', 70.5], ['Drinks', 122.75], ['Supplies', 167], ['Apparel', 221]];

  C.lessons.push({
    week: 9, unit: 2, id: 'week-9',
    title: 'Charts that tell the truth',
    summary: 'Matching the chart to the question (comparison, trend, part-to-whole, relationship), the parts of a chart, design rules that make charts readable, and the common ways charts mislead.',
    objectives: [
      'Choose a chart type for comparison, trend over time, part-to-whole and relationship questions.',
      'Name the parts of a chart: title, axes, labels, legend, data series.',
      'Apply design rules: one message per chart, a title that states the finding, sorted bars, minimal clutter.',
      'Identify at least three ways a chart can mislead and fix them.',
    ],
    sections: [
      { heading: 'Start with the question', html: `
        <table class="notes-table">
          <thead><tr><th>The question is about…</th><th>Use</th><th>Example</th></tr></thead>
          <tbody>
            <tr><td><strong>Comparison</strong> between categories</td><td>Column or bar chart (bar when labels are long)</td><td>Revenue by category</td></tr>
            <tr><td><strong>Trend</strong> over time</td><td>Line chart (time always runs left to right)</td><td>Revenue by month</td></tr>
            <tr><td><strong>Part-to-whole</strong></td><td>Pie (2–5 slices only) or a stacked/100% bar</td><td>Card vs cash share of orders</td></tr>
            <tr><td><strong>Relationship</strong> between two measures</td><td>Scatter plot</td><td>Hours studied vs test score</td></tr>
            <tr><td>One important number</td><td>Not a chart: a big number ("card") with context</td><td>Total revenue: $591.50</td></tr>
          </tbody>
        </table>` },
      { heading: 'Anatomy of a chart', html: `
        <ul>
          <li><strong>Title:</strong> best when it states the takeaway ("Apparel brings in the most revenue"), not just the topic ("Revenue by Category").</li>
          <li><strong>Axes:</strong> the category axis (x for column charts) and the value axis, with units.</li>
          <li><strong>Data series:</strong> one set of values; each series gets one color.</li>
          <li><strong>Legend:</strong> only needed when there is more than one series.</li>
          <li><strong>Data labels:</strong> use sparingly, for the numbers that matter.</li>
        </ul>
        <p>In Excel, select a summary table (labels in the first column, values next to it) and choose <strong>Insert › Recommended Charts</strong>. In the practice workbook, use <strong>Chart</strong> on the ribbon. Charts update automatically when the data changes.</p>` },
      { heading: 'Design rules', html: `
        <ul>
          <li><strong>Chart a summary, not raw rows.</strong> 60 order rows make an unreadable chart; 4 category totals make a clear one.</li>
          <li><strong>Sort bars</strong> largest to smallest unless the categories have a natural order (months, sizes).</li>
          <li><strong>Start bar and column axes at zero.</strong> Bar length is the message.</li>
          <li><strong>Remove clutter:</strong> 3-D effects, heavy gridlines, shadows and decorative backgrounds add nothing.</li>
          <li><strong>Use color with purpose:</strong> one color for everything, plus one highlight color for the bar you are talking about.</li>
          <li><strong>Make it accessible:</strong> do not rely on red vs green alone, and label things directly when you can.</li>
        </ul>` },
      { heading: 'How charts mislead', html: `
        <ol>
          <li><strong>Truncated axis:</strong> a column chart whose axis starts at 90 makes 91 vs 95 look like a 5× difference. Practice 9B lets you see this happen.</li>
          <li><strong>Cherry-picked time range:</strong> showing only the three months that went up.</li>
          <li><strong>Dual axes:</strong> two different scales on one chart can make any two lines look related.</li>
          <li><strong>Too many pie slices</strong> or a pie that does not add up to 100%.</li>
          <li><strong>3-D and area tricks:</strong> perspective makes front slices look bigger; doubling an icon's width <em>and</em> height shows 4× the area for 2× the value.</li>
          <li><strong>Correlation shown as causation:</strong> two lines moving together does not prove one causes the other.</li>
        </ol>
        <div class="callout">Line charts may start above zero when you care about small changes in a trend (like temperature), as long as the axis is clearly labeled. Bar and column charts should not.</div>` },
    ],
    keyTerms: [
      ['Column / bar chart', 'Compares values across categories.'],
      ['Line chart', 'Shows change over time.'],
      ['Pie chart', 'Shows parts of a whole (few slices only).'],
      ['Scatter plot', 'Shows the relationship between two numeric measures.'],
      ['Truncated axis', 'A value axis that does not start at zero, exaggerating differences.'],
      ['Data series', 'A set of related values plotted together.'],
    ],
    practice: [
      {
        id: 'p9a', title: 'Workbook 9A · Pick the right chart',
        intro: 'Two summaries are ready. Select each table (including its header) and insert the chart that fits the question. Give each chart a title that states the finding.',
        tools: ['format', 'bold', 'fill', 'sort', 'chart', 'formulas', 'download', 'reset'],
        workbook: {
          sheets: [{
            name: 'Summary',
            rows: [...monthly.map((r) => [...r, null, null, null]), [null], [null], ...byCategory.map((r) => [...r])],
            formats: { 'B2:B5': 'currency', 'B9:B12': 'currency' },
            widths: { A: 90, B: 90 },
            lock: ['A1:B5', 'A8:B8'],
            cells: { D1: { input: 'Q1: How did revenue change from month to month? (chart A1:B5)', locked: true }, D8: { input: 'Q2: Which category brings in the most revenue? (chart A8:B12, sorted)', locked: true } },
          }],
        },
        questions: [
          { id: 'q1', type: 'chart', sheet: 'Summary', chartType: 'line', solution: { type: 'line', range: 'A1:B5', title: 'Revenue dipped in Feb–Mar, then hit a high in April' }, requireTitle: true, points: 3, prompt: 'A line chart of A1:B5 with a title.' },
          { id: 'q2', type: 'sorted', sheet: 'Summary', range: 'A8:B12', column: 'B', order: 'desc', points: 2, prompt: 'Sort the category table by Revenue, largest first.' },
          { id: 'q3', type: 'chart', sheet: 'Summary', chartType: ['column', 'bar'], solution: { type: 'column', range: 'A8:B12', title: 'Apparel brings in the most revenue' }, requireTitle: true, axisFromZero: true, points: 3, prompt: 'A column or bar chart of A8:B12, axis starting at 0, with a title.' },
          { id: 'q4', type: 'choice', prompt: 'Which chart best shows how two numeric measures relate (e.g. hours studied vs score)?', options: ['Pie', 'Line', 'Scatter', 'Stacked column'], answer: 2 },
        ],
      },
      {
        id: 'p9b', title: 'Workbook 9B · Same data, different story',
        intro: 'Insert a column chart of A1:B4. Then edit the chart and set "Value axis starts at" to 90. Same numbers, very different impression. Set it back to 0 (blank) before you check your work.',
        tools: ['chart', 'download', 'reset'],
        workbook: { sheets: [{ name: 'Scores', rows: [['Class', 'Average Score'], ['Period 1', 91], ['Period 2', 93], ['Period 3', 95]], lock: true, widths: { B: 110 } }] },
        questions: [
          { id: 'q1', type: 'chart', sheet: 'Scores', chartType: ['column', 'bar'], solution: { type: 'column', range: 'A1:B4', title: 'Average score by class period' }, axisFromZero: true, points: 2, prompt: 'A column chart of A1:B4 whose axis starts at 0.' },
          { id: 'q2', type: 'number', prompt: 'With the axis starting at 90, Period 3\'s column looks about how many times taller than Period 1\'s? (95−90) ÷ (91−90) = ?', answer: 5 },
          { id: 'q3', type: 'number', prompt: 'In reality, Period 3\'s average is what percent higher than Period 1\'s? Round to one decimal (e.g. 4.4).', answer: 4.4, tolerance: 0.05 },
        ],
      },
    ],
    assignment: {
      id: 'hw-9', title: 'Chart story: what happened at the store?', points: 25,
      html: `<p>Build three honest charts that answer three different kinds of questions about the store, each with a takeaway title.</p>
        <ol><li>On the <code>Charts</code> sheet, finish the summary tables using SUMIFS / COUNTIF on the Sales sheet.</li>
        <li>Insert a <strong>line</strong> chart for revenue by month (A3:B7), a <strong>column or bar</strong> chart for revenue by category (D3:E7, sorted largest first after you fill it), and a <strong>pie</strong> chart for orders by payment type (G3:H5).</li>
        <li>Every chart needs a title that states a finding. Column/bar axes start at 0.</li>
        <li>In <code>A10</code>, write two sentences: what is the most important thing these charts show?</li></ol>`,
      tools: 'analyze',
      workbook: {
        sheets: [
          (() => {
            const s = salesWithRevenue('Sales');
            s.rows[0] = [...s.rows[0], 'Month'];
            for (let r = 2; r <= N; r++) s.cells['K' + r] = { input: '=TEXT(B' + r + ',"mmm")', locked: true };
            s.lock = 'A1:K' + N;
            return s;
          })(),
          {
            name: 'Charts',
            rows: [
              ['Store chart story'], [null],
              ['Month', 'Revenue', null, 'Category', 'Revenue', null, 'Payment', 'Orders'],
              ['Jan', null, null, 'Snacks', null, null, 'Card'],
              ['Feb', null, null, 'Drinks', null, null, 'Cash'],
              ['Mar', null, null, 'Supplies'],
              ['Apr', null, null, 'Apparel'],
              [null], ['Most important takeaway'],
            ],
            formats: { 'B4:B7': 'currency', 'E4:E7': 'currency' },
            widths: { A: 90, D: 90, G: 80 },
            lock: ['A1:A9', 'A3:H3', 'G4:G5'],
            styles: { 'B4:B7': { fill: 'answer' }, 'E4:E7': { fill: 'answer' }, 'H4:H5': { fill: 'answer' }, A10: { fill: 'answer' } },
          },
        ],
      },
      model: (wb) => {
        // fill the category table, then sort it (sorting needs the values present)
        for (let r = 4; r <= 7; r++) wb.setInput('Charts', 'E' + r, '=SUMIF(Sales!$E$2:$E$61,D' + r + ',Sales!$J$2:$J$61)', { force: true });
        wb.recalc();
        wb.sortRange('Charts', 'D3:E7', 1, false, true);
      },
      questions: [
        { id: 'a1', type: 'range', sheet: 'Charts', range: 'B4:B7', solution: '=SUMIF(Sales!$K$2:$K$61,A4,Sales!$J$2:$J$61)', requireFormula: true, points: 4, prompt: 'Revenue by month (B4:B7).' },
        { id: 'a2', type: 'range', sheet: 'Charts', range: 'E4:E7', expect: [[221], [167], [122.75], [70.5]], solution: null, requireFormula: true, points: 4, prompt: 'Revenue by category (E4:E7), sorted largest first.' },
        { id: 'a3', type: 'range', sheet: 'Charts', range: 'H4:H5', solution: '=COUNTIF(Sales!$H$2:$H$61,G4)', requireFormula: true, points: 2, prompt: 'Orders by payment type (H4:H5).' },
        { id: 'a4', type: 'chart', sheet: 'Charts', chartType: 'line', solution: { type: 'line', range: 'A3:B7', title: 'Revenue dipped in Feb–Mar and peaked in April' }, requireTitle: true, points: 4, prompt: 'Line chart of revenue by month, titled.' },
        { id: 'a5', type: 'chart', sheet: 'Charts', chartType: ['column', 'bar'], solution: { type: 'bar', range: 'D3:E7', title: 'Apparel earns the most despite few orders' }, requireTitle: true, axisFromZero: true, points: 4, prompt: 'Column or bar chart of revenue by category, titled, axis from 0.' },
        { id: 'a6', type: 'chart', sheet: 'Charts', chartType: 'pie', solution: { type: 'pie', range: 'G3:H5', title: 'Slightly more orders are paid in cash' }, requireTitle: true, points: 4, prompt: 'Pie chart of orders by payment, titled.' },
        { id: 'a7', type: 'filled', sheet: 'Charts', range: 'A10', min: 1, points: 3, prompt: 'Takeaway written in A10.' },
      ],
    },
    quiz: 'quiz-9',
  });

  // ---------------------------------------------------------------- Week 10
  C.lessons.push({
    week: 10, unit: 2, id: 'week-10',
    title: 'Pivot tables and Unit 2 review',
    summary: 'Summarizing data by dimension with PivotTables: rows, columns, values and filters; changing the aggregation; grouping dates to change granularity; slicers and refresh. Then a review of Unit 2 before Friday\'s exam.',
    objectives: [
      'Build a PivotTable that aggregates a measure by a dimension.',
      'Change the aggregation (sum, count, average, max) and explain what each answers.',
      'Change granularity by grouping dates into months or quarters.',
      'Add a second dimension in Columns, filter with slicers, and refresh after the data changes.',
      'Explain why a pivot and a SUMIFS table give the same answer.',
    ],
    sections: [
      { heading: 'What a PivotTable does', html: `
        <p>In week 6 you built a category summary with one SUMIFS per cell. A <strong>PivotTable</strong> builds the same summary by dragging fields, without writing formulas. It <em>groups</em> rows by a dimension and <em>aggregates</em> a measure for each group.</p>
        <p>To create one in Excel: click inside your data (ideally an Excel Table), then <strong>Insert › PivotTable</strong>, and choose a new worksheet. In the practice workbook, use <strong>PivotTable</strong> on the ribbon.</p>` },
      { heading: 'The four areas', html: `
        <table class="notes-table">
          <thead><tr><th>Area</th><th>What goes there</th><th>Example</th></tr></thead>
          <tbody>
            <tr><td><strong>Rows</strong></td><td>A dimension; one row per value</td><td>Category</td></tr>
            <tr><td><strong>Columns</strong></td><td>A second dimension across the top</td><td>Payment</td></tr>
            <tr><td><strong>Values</strong></td><td>A measure and how to aggregate it</td><td>Sum of Revenue</td></tr>
            <tr><td><strong>Filters / Slicers</strong></td><td>Limit which rows are included</td><td>Only Cashier = Sam</td></tr>
          </tbody>
        </table>
        <p>Put Category in Rows and Revenue in Values and you have the week 6 table, including a Grand Total, in a few seconds.</p>` },
      { heading: 'Aggregation: what question are you asking?', html: `
        <ul>
          <li><strong>Sum</strong> of Revenue: how much money in total?</li>
          <li><strong>Count</strong> of OrderID: how many orders?</li>
          <li><strong>Average</strong> of Revenue: how big is a typical order?</li>
          <li><strong>Max</strong> / <strong>Min</strong>: the largest / smallest single order.</li>
        </ul>
        <p>Excel defaults to Sum for numbers and Count for text. Always read the header ("Sum of Revenue", "Count of OrderID") to make sure the pivot answers your question. Summing an ID column is a classic mistake: the total of order numbers means nothing.</p>` },
      { heading: 'Granularity: grouping dates', html: `
        <p><strong>Granularity</strong> is the level of detail. Daily sales are fine-grained; monthly totals are coarse. Put Date in Rows and Excel lists every day. Group it (right-click › Group › Months, or choose Months in the practice dialog) to see one row per month. Coarser granularity hides noise and shows the trend; finer granularity lets you investigate a specific day.</p>` },
      { heading: 'Slicers and refresh', html: `
        <p>A <strong>slicer</strong> is a set of buttons that filters a PivotTable (and any chart built on it) with one click: <em>PivotTable Analyze › Insert Slicer</em>. Slicers make filters visible, so the reader knows what they are looking at. They are also the main way people filter in Power BI.</p>
        <p>A PivotTable is a <strong>snapshot</strong>. When the source data changes, click <strong>Refresh</strong> (Data › Refresh All). If your source is an Excel Table, new rows are included automatically on refresh.</p>` },
      { heading: 'Unit 2 review checklist', html: `
        <p>Exam 2 is Friday and covers weeks 6–10 only. You should be able to:</p>
        <ul class="checklist-review">
          <li>Turn a vague question into a measurable one; tell measures from dimensions and metrics from KPIs.</li>
          <li>Write IF, IFS, AND/OR, COUNTIF(S), SUMIF(S), AVERAGEIF(S) and MAXIFS with text, number, wildcard and cell-reference criteria.</li>
          <li>Explain primary keys, foreign keys, one-to-many relationships, referential integrity and orphan records.</li>
          <li>Join tables with XLOOKUP (and read VLOOKUP with FALSE and INDEX/MATCH); handle #N/A.</li>
          <li>Clean text with TRIM, PROPER, UPPER, LOWER, SUBSTITUTE, LEFT/RIGHT/MID/FIND; extract date parts; remove duplicates.</li>
          <li>Pick a chart for comparison, trend, part-to-whole or relationship; spot a truncated axis and other tricks.</li>
          <li>Build a PivotTable, change aggregation and date grouping, and explain slicers and refresh.</li>
        </ul>` },
    ],
    keyTerms: [
      ['PivotTable', 'An interactive summary that groups and aggregates data.'],
      ['Aggregation', 'Combining many values into one: sum, count, average, min, max.'],
      ['Granularity', 'The level of detail of the data (day, month, year).'],
      ['Slicer', 'A visual button filter for PivotTables and reports.'],
      ['Refresh', 'Recalculate a PivotTable or query from the current source data.'],
    ],
    practice: [
      {
        id: 'p10a', title: 'Workbook 10A · The week 6 summary, rebuilt with pivots',
        intro: 'Click PivotTable. Build each pivot from the Sales data (A1:J61), then answer the questions on the Answers sheet using what the pivots show.',
        tools: 'analyze',
        workbook: {
          sheets: [
            salesWithRevenue('Sales'),
            { name: 'Answers', rows: [['Question', 'Answer'], ['Revenue from Supplies'], ['Number of orders handled by Avery'], ['Revenue in April']], formats: { B2: 'currency', B4: 'currency' }, lock: 'A1:A4', widths: { A: 300, B: 100 }, styles: { 'B2:B4': { fill: 'answer' } } },
          ],
        },
        questions: [
          { id: 'q1', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:J61', rows: 'Category', values: 'Revenue', agg: 'sum' }, points: 2, prompt: 'Pivot 1: Category in Rows, Sum of Revenue in Values.' },
          { id: 'q2', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:J61', rows: 'Cashier', values: 'OrderID', agg: 'count' }, points: 2, prompt: 'Pivot 2: Cashier in Rows, Count of OrderID in Values.' },
          { id: 'q3', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:J61', rows: 'Date', values: 'Revenue', agg: 'sum', rowsGroup: 'month' }, points: 2, prompt: 'Pivot 3: Date grouped by month in Rows, Sum of Revenue in Values.' },
          { id: 'q4', type: 'cell', sheet: 'Answers', cell: 'B2', solution: '=SUMIF(Sales!E2:E61,"Supplies",Sales!J2:J61)', tolerance: 0.005, prompt: 'Revenue from Supplies (type the number from your pivot).' },
          { id: 'q5', type: 'cell', sheet: 'Answers', cell: 'B3', solution: '=COUNTIF(Sales!I2:I61,"Avery")', prompt: 'Orders handled by Avery.' },
          { id: 'q6', type: 'cell', sheet: 'Answers', cell: 'B4', solution: '=SUMIFS(Sales!J2:J61,Sales!B2:B61,">="&DATE(2026,4,1),Sales!B2:B61,"<"&DATE(2026,5,1))', tolerance: 0.005, prompt: 'Revenue in April.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-10', title: 'Pivot analysis memo', points: 25,
      html: `<p>The principal wants to know whether to keep the store open on Mondays and whether apparel is worth stocking. Use PivotTables to find out.</p>
        <ol><li>Build a pivot of <strong>Sum of Revenue by Category</strong> with <strong>Payment in Columns</strong>.</li>
        <li>Build a pivot of <strong>Count of OrderID by Weekday</strong> (column K).</li>
        <li>Fill the Memo sheet: the numbers in B2:B4 and two recommendations in B6:B7.</li></ol>`,
      tools: 'analyze',
      workbook: {
        sheets: [
          (() => {
            const s = salesWithRevenue('Sales');
            s.rows[0] = [...s.rows[0], 'Weekday'];
            for (let r = 2; r <= N; r++) s.cells['K' + r] = { input: '=TEXT(B' + r + ',"ddd")', locked: true };
            s.lock = 'A1:K' + N;
            return s;
          })(),
          {
            name: 'Memo',
            rows: [['Finding', 'Value'], ['Apparel revenue paid by Card'], ['Orders on Mondays'], ['Average revenue per Apparel order'], [null], ['Recommendation about Mondays'], ['Recommendation about Apparel']],
            formats: { B2: 'currency', B4: 'currency' },
            lock: 'A1:A7', widths: { A: 300, B: 440 },
            styles: { 'B2:B4': { fill: 'answer' }, 'B6:B7': { fill: 'answer' } },
          },
        ],
      },
      questions: [
        { id: 'a1', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:K61', rows: 'Category', columns: 'Payment', values: 'Revenue', agg: 'sum' }, points: 6, prompt: 'Pivot: Sum of Revenue by Category × Payment.' },
        { id: 'a2', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:K61', rows: 'Weekday', values: 'OrderID', agg: 'count' }, points: 5, prompt: 'Pivot: Count of OrderID by Weekday.' },
        { id: 'a3', type: 'cell', sheet: 'Memo', cell: 'B2', solution: '=SUMIFS(Sales!J2:J61,Sales!E2:E61,"Apparel",Sales!H2:H61,"Card")', tolerance: 0.005, points: 3, prompt: 'Apparel revenue paid by Card.' },
        { id: 'a4', type: 'cell', sheet: 'Memo', cell: 'B3', solution: '=COUNTIF(Sales!K2:K61,"Mon")', points: 3, prompt: 'Orders on Mondays.' },
        { id: 'a5', type: 'cell', sheet: 'Memo', cell: 'B4', solution: '=AVERAGEIF(Sales!E2:E61,"Apparel",Sales!J2:J61)', tolerance: 0.005, points: 3, prompt: 'Average revenue per Apparel order.' },
        { id: 'a6', type: 'filled', sheet: 'Memo', range: 'B6:B7', min: 2, points: 5, prompt: 'Both recommendations written.' },
      ],
    },
    quiz: 'quiz-10',
    exam: 'exam-2',
  });

  if (isNode) module.exports = C;
})(typeof globalThis !== 'undefined' ? globalThis : this);
