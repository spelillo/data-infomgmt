/*
 * unit1.js — Unit 1: Data foundations and Excel basics (weeks 1–5).
 */
(function (root) {
  'use strict';
  const isNode = typeof module === 'object' && module.exports;
  const C = isNode ? require('./course.js') : root.DIM.Content;
  const D = C.data;

  // ---------------------------------------------------------------- Week 1
  C.lessons.push({
    week: 1, unit: 1, id: 'week-1',
    title: 'What data is and why it matters',
    summary: 'Data, information and knowledge; how organizations turn data into decisions; a first tour of Excel.',
    objectives: [
      'Define data, information and knowledge, and give an example of each.',
      'Describe the data-to-decision cycle an organization follows.',
      'Name the parts of the Excel window: workbook, worksheet, cell, cell address, ribbon, Name Box, formula bar, sheet tabs, status bar.',
      'Move around a worksheet with the keyboard and enter data.',
    ],
    sections: [
      { heading: 'Data is recorded observation', html: `
        <p><strong>Data</strong> is any fact or observation that has been recorded so it can be used later. The time your alarm went off, the price of a sports drink, the number of steps on your phone, the score of last night's game: each is a data point.</p>
        <p>Data on its own is raw. A list like <em>7, 6, 8, 6.5, 7</em> means nothing until you know it is "hours of sleep for five students last night". The label and the context are part of what makes data useful.</p>
        <div class="callout"><strong>Think about it:</strong> before you got to class today, how many data points did apps, cameras, card readers and your school system record about you? Most students estimate 10; the real number is usually in the hundreds.</div>` },
      { heading: 'From data to information to knowledge', html: `
        <p>A useful way to think about data is as the bottom of a ladder. Each step up adds context and meaning.</p>
        <table class="notes-table">
          <thead><tr><th>Level</th><th>What it is</th><th>Class survey example</th></tr></thead>
          <tbody>
            <tr><td><strong>Data</strong></td><td>Raw recorded facts</td><td>"Pizza", "Tacos", "Pizza", "Salad"… typed one per student</td></tr>
            <tr><td><strong>Information</strong></td><td>Data that has been organized or summarized so it answers a question</td><td>5 of 15 students picked pizza; it is the most popular lunch</td></tr>
            <tr><td><strong>Knowledge</strong></td><td>Understanding that comes from interpreting information with experience</td><td>Pizza days sell out, so the cafeteria should order more on those days</td></tr>
            <tr><td><strong>Decision / action</strong></td><td>What someone does because of the knowledge</td><td>The cafeteria orders 20% more pizza and tracks whether waste goes down</td></tr>
          </tbody>
        </table>
        <p>This is sometimes called the <strong>DIKW pyramid</strong> (Data, Information, Knowledge, Wisdom). The point to remember: <em>organizing and summarizing data is what turns it into information</em>, and that is exactly what spreadsheets are for.</p>` },
      { heading: 'Why data matters: how organizations decide', html: `
        <p>Every organization makes decisions under uncertainty. Data reduces that uncertainty. Most data-driven decisions follow the same cycle:</p>
        <ol class="cycle">
          <li><strong>Ask</strong> a specific question. ("Which snacks should the school store stop selling?")</li>
          <li><strong>Collect</strong> the data that could answer it. (Every sale from the register.)</li>
          <li><strong>Organize</strong> it so it can be analyzed. (One row per sale, consistent item names.)</li>
          <li><strong>Analyze</strong> it. (Total units and revenue per item.)</li>
          <li><strong>Decide</strong> and act. (Drop the two slowest sellers.)</li>
          <li><strong>Measure</strong> the result and start again. (Did total revenue go up?)</li>
        </ol>
        <p>Real examples you can discuss: a grocery chain predicting how much milk to stock before a storm; a district planning bus routes from student addresses; a basketball coach choosing lineups from shooting percentages; a streaming service picking what to recommend next. In each case, poor-quality data leads to poor decisions, which is why the rest of this course spends so much time on getting data right.</p>` },
      { heading: 'Where data lives: rows, columns and tables', html: `
        <p>Most business data is stored in <strong>tables</strong>. A table has:</p>
        <ul>
          <li><strong>Columns</strong> (also called <em>fields</em> or <em>variables</em>): one kind of fact, like Grade or Favorite Lunch.</li>
          <li><strong>Rows</strong> (also called <em>records</em> or <em>observations</em>): one thing being described, like one student.</li>
          <li>A <strong>header row</strong> at the top that names each column.</li>
        </ul>
        <p>Spreadsheets like Excel and Google Sheets, databases like SQL Server, and business apps all store data this way. If you can read a table, you can read most of the world's data.</p>` },
      { heading: 'A tour of Excel', html: `
        <p>Open the practice workbook below while you read. It behaves like Excel: click a cell, type, press <kbd>Enter</kbd>.</p>
        <table class="notes-table">
          <thead><tr><th>Part</th><th>What it does</th></tr></thead>
          <tbody>
            <tr><td><strong>Workbook</strong></td><td>The whole file (.xlsx). It can hold many worksheets.</td></tr>
            <tr><td><strong>Worksheet</strong> (sheet)</td><td>One grid of cells. Switch sheets with the tabs at the bottom.</td></tr>
            <tr><td><strong>Cell</strong></td><td>One box in the grid. It holds one value or one formula.</td></tr>
            <tr><td><strong>Cell address</strong></td><td>Column letter + row number, like <code>C4</code>. Columns go A–Z, then AA, AB…</td></tr>
            <tr><td><strong>Range</strong></td><td>A block of cells written with a colon: <code>A2:F16</code> means "A2 through F16".</td></tr>
            <tr><td><strong>Name Box</strong></td><td>Top left. Shows the active cell; type an address there and press Enter to jump to it.</td></tr>
            <tr><td><strong>Formula bar</strong></td><td>Shows what is really in the cell (a value or a formula). You can edit here too.</td></tr>
            <tr><td><strong>Ribbon</strong></td><td>The toolbar of commands (Home, Insert, Data…). The practice workbook has a simplified ribbon.</td></tr>
            <tr><td><strong>Status bar</strong></td><td>Bottom right. Select several numbers and it shows their Average, Count and Sum, with no formula needed.</td></tr>
          </tbody>
        </table>` },
      { heading: 'Moving around and entering data', html: `
        <table class="notes-table">
          <thead><tr><th>Keys</th><th>What happens</th></tr></thead>
          <tbody>
            <tr><td><kbd>Arrow keys</kbd></td><td>Move one cell</td></tr>
            <tr><td><kbd>Enter</kbd> / <kbd>Shift</kbd>+<kbd>Enter</kbd></td><td>Confirm and move down / up</td></tr>
            <tr><td><kbd>Tab</kbd> / <kbd>Shift</kbd>+<kbd>Tab</kbd></td><td>Confirm and move right / left: the fastest way to type a row of data</td></tr>
            <tr><td><kbd>Ctrl</kbd>+<kbd>Arrow</kbd></td><td>Jump to the edge of the data (great for big tables)</td></tr>
            <tr><td><kbd>Ctrl</kbd>+<kbd>Home</kbd></td><td>Go to A1</td></tr>
            <tr><td><kbd>Shift</kbd>+<kbd>Arrow</kbd></td><td>Select a range</td></tr>
            <tr><td><kbd>F2</kbd></td><td>Edit the active cell without erasing it</td></tr>
            <tr><td><kbd>Esc</kbd></td><td>Cancel what you are typing</td></tr>
            <tr><td><kbd>Ctrl</kbd>+<kbd>Z</kbd> / <kbd>Ctrl</kbd>+<kbd>Y</kbd></td><td>Undo / Redo</td></tr>
          </tbody>
        </table>
        <div class="callout tip"><strong>Habit to build now:</strong> type one record per row and one kind of fact per column, with a single header row. Everything you learn later (sorting, formulas, pivot tables, Power BI) depends on it.</div>` },
    ],
    keyTerms: [
      ['Data', 'Recorded facts or observations.'],
      ['Information', 'Data organized or summarized so it answers a question.'],
      ['Knowledge', 'Understanding gained by interpreting information.'],
      ['Record (row)', 'One instance of the thing being described.'],
      ['Field (column)', 'One attribute recorded for every record.'],
      ['Cell address', 'The column letter and row number of a cell, e.g. B7.'],
      ['Range', 'A rectangular group of cells, e.g. A2:F16.'],
    ],
    practice: [
      {
        id: 'p1a', title: 'Workbook 1A · Class survey: from data to information',
        intro: 'The survey data is locked so you cannot damage it. Answer in the green cells. Use the status bar, the Name Box and your own counting. No formulas needed yet.',
        tools: 'basic',
        workbook: {
          sheets: [{
            name: 'Survey',
            rows: D.survey.map((r, i) => (i === 0 ? [...r, null, 'Question', 'Your answer'] : [...r])),
            cells: {
              H2: 'How many students chose Pizza?', H3: 'How many students have a job?', H4: 'Average hours of sleep (select D2:D16 and read the status bar)', H5: 'One decision a school could make from this data',
            },
            lock: true,
            widths: { C: 110, D: 105, E: 80, F: 120, H: 330, I: 260 },
            styles: { 'I2:I5': { fill: 'answer' } },
          }],
        },
        questions: [
          { id: 'q1', type: 'cell', sheet: 'Survey', cell: 'I2', solution: '=COUNTIF(C2:C16,"Pizza")', prompt: 'In I2, type how many students chose Pizza.', hint: 'Count the word Pizza in column C.' },
          { id: 'q2', type: 'cell', sheet: 'Survey', cell: 'I3', solution: '=COUNTIF(E2:E16,"Yes")', prompt: 'In I3, type how many students have a job.' },
          { id: 'q3', type: 'cell', sheet: 'Survey', cell: 'I4', solution: '=AVERAGE(D2:D16)', tolerance: 0.05, prompt: 'Select D2:D16, read the Average in the status bar, and type it in I4.' },
          { id: 'q4', type: 'filled', sheet: 'Survey', range: 'I5', min: 1, prompt: 'In I5, write one decision someone could make from this information.' },
          { id: 'q5', type: 'filled', sheet: 'Survey', range: 'K25', min: 1, prompt: 'Click the Name Box, type K25, press Enter, and type your initials there.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-1', title: 'Data diary', points: 20,
      html: `<p>For one full day, record at least <strong>20 things you did</strong> as data. Each row is one activity; each column is one fact about it.</p>
        <ol><li>Open the starter workbook and keep the header row as is.</li>
        <li>Fill at least 20 rows. Use a number (no units) in the Minutes column.</li>
        <li>In <code>H2</code>, write one piece of <em>information</em> your data shows (for example "I spent 95 minutes on my phone, more than on homework").</li>
        <li>Turn it in on the Assignments page (upload your .xlsx or submit the web workbook).</li></ol>`,
      workbook: { sheets: [{ name: 'Diary', rows: [['Start Time', 'Activity', 'Category', 'Minutes', 'With Whom', null, null, 'One piece of information from my data']], widths: { A: 85, B: 160, C: 100, D: 70, E: 100, H: 380 }, lock: 'A1:H1' }] },
      questions: [
        { id: 'a1', type: 'filled', sheet: 'Diary', range: 'B2:B60', min: 20, points: 8, prompt: 'At least 20 activities recorded.' },
        { id: 'a2', type: 'filled', sheet: 'Diary', range: 'D2:D60', min: 20, points: 4, prompt: 'Minutes recorded for each activity.' },
        { id: 'a3', type: 'filled', sheet: 'Diary', range: 'C2:C60', min: 20, points: 4, prompt: 'Every activity has a category.' },
        { id: 'a4', type: 'filled', sheet: 'Diary', range: 'H2', min: 1, points: 4, prompt: 'One piece of information written in H2.' },
      ],
      rubric: [['Minutes stored as plain numbers (no "min")', 'Teacher check'], ['Categories are consistent (same spelling each time)', 'Teacher check']],
    },
    quiz: null,
  });

  // ---------------------------------------------------------------- Week 2
  const typeRows = [
    ['Field', 'Example value', 'Type (Nominal / Ordinal / Discrete / Continuous)'],
    ['Favorite lunch', 'Pizza', null],
    ['T-shirt size', 'Medium', null],
    ['Number of siblings', 2, null],
    ['Height (inches)', 65.5, null],
    ['Class rank (1st, 2nd, 3rd)', '2nd', null],
    ['Home state', 'Ohio', null],
    ['Minutes to get to school', 17.25, null],
    ['Number of absences', 3, null],
  ];
  const typeAnswers = ['Nominal', 'Ordinal', 'Discrete', 'Continuous', 'Ordinal', 'Nominal', 'Continuous', 'Discrete'];

  C.lessons.push({
    week: 2, unit: 1, id: 'week-2',
    title: 'Data types and tidy tables',
    summary: 'Qualitative vs quantitative data, the four levels of measurement, how Excel stores types, and the rules of a tidy table.',
    objectives: [
      'Classify a field as nominal, ordinal, discrete or continuous.',
      'Tell structured, semi-structured and unstructured data apart.',
      'Explain how Excel stores numbers, text, dates and TRUE/FALSE, and spot numbers stored as text.',
      'Apply number formats without changing the stored value.',
      'Build a tidy table: one row per record, one column per field, one value per cell.',
    ],
    sections: [
      { heading: 'Qualitative and quantitative data', html: `
        <p><strong>Qualitative</strong> (categorical) data describes a quality: a name, a color, a category. <strong>Quantitative</strong> (numerical) data measures or counts something, so arithmetic on it makes sense.</p>
        <p>Each splits in two, giving the four <strong>levels of measurement</strong>:</p>
        <table class="notes-table">
          <thead><tr><th>Type</th><th>Meaning</th><th>Examples</th><th>What you can do with it</th></tr></thead>
          <tbody>
            <tr><td><strong>Nominal</strong></td><td>Categories with no order</td><td>Favorite lunch, home state, payment method</td><td>Count, find the most common</td></tr>
            <tr><td><strong>Ordinal</strong></td><td>Categories with a meaningful order</td><td>T-shirt size, class rank, "Strongly agree…disagree"</td><td>Count, sort, find the middle</td></tr>
            <tr><td><strong>Discrete</strong></td><td>Counts: whole numbers</td><td>Siblings, absences, items sold</td><td>All math; no fractions of a unit</td></tr>
            <tr><td><strong>Continuous</strong></td><td>Measurements that can take any value</td><td>Height, time, temperature, price</td><td>All math, averages, decimals</td></tr>
          </tbody>
        </table>
        <div class="callout"><strong>Trap:</strong> a number is not always quantitative. A zip code, jersey number or student ID is made of digits, but adding two zip codes is meaningless. Those are <em>nominal</em> and should be stored as <em>text</em>.</div>` },
      { heading: 'Structured, semi-structured and unstructured data', html: `
        <ul>
          <li><strong>Structured:</strong> fits neatly in rows and columns with a fixed set of fields. Spreadsheets, database tables, register sales.</li>
          <li><strong>Semi-structured:</strong> has labels but not a fixed table shape. JSON from an app, emails (To/From/Subject plus a free-text body), XML.</li>
          <li><strong>Unstructured:</strong> no predefined fields. Photos, videos, voice memos, essays, social media posts. Estimates put 80% or more of the world's data here.</li>
        </ul>
        <p>This course focuses on structured data, because that is where decisions are usually made, and much of a data professional's job is turning the other two kinds into structured tables.</p>` },
      { heading: 'How Excel stores types', html: `
        <p>Every cell holds exactly one of: a <strong>number</strong>, <strong>text</strong>, a <strong>logical value</strong> (TRUE/FALSE), an <strong>error</strong> (like <code>#DIV/0!</code>), or nothing. Excel decides the type when you press Enter.</p>
        <ul>
          <li>Numbers line up on the <strong>right</strong> of the cell; text lines up on the <strong>left</strong>. This is the fastest way to spot a "number" that is really text.</li>
          <li><strong>Dates are numbers.</strong> Excel stores 3/15/2026 as <code>46096</code>, the number of days since January 1, 1900, and only <em>displays</em> it as a date. That is why you can subtract two dates to get days between them.</li>
          <li>Typing <code>00501</code> gives the number 501: the leading zeros are gone. To keep them, type an apostrophe first (<code>'00501</code>) or format the cells as Text before typing.</li>
          <li>A number stored as text will not add up: <code>SUM</code> skips it. The practice workbook shows a hint above the grid when it finds one.</li>
        </ul>
        <div class="callout tip"><strong>Try it:</strong> in Workbook 2B, type <code>3/15/2026</code> in an empty cell, then change its format to General. You will see 46096.</div>` },
      { heading: 'Formats change how a value looks, not what it is', html: `
        <p>A <strong>number format</strong> is a display rule. The value <code>0.256</code> can appear as <code>0.256</code>, <code>26%</code>, <code>25.6%</code> or <code>$0.26</code>. Calculations always use the stored value. Check the formula bar to see the truth.</p>
        <p>Common formats: General, Number, Currency, Percent, Short Date, Long Date and Text. In the practice workbook, select cells and use the format drop-down on the ribbon.</p>` },
      { heading: 'Tidy tables', html: `
        <p>A table is <strong>tidy</strong> when:</p>
        <ol>
          <li>Each <strong>column</strong> is one variable (field).</li>
          <li>Each <strong>row</strong> is one observation (record).</li>
          <li>Each <strong>cell</strong> holds one value.</li>
        </ol>
        <p>Plus a few spreadsheet rules that save you hours later: one header row with unique names, no merged cells, no blank rows or columns inside the data, no totals mixed into the data, and units in the header (<code>Minutes</code>), not in the cells (<code>15 min</code>).</p>
        <p>Many reports are built for people to read, not for computers to analyze. A "wide" report with one column per month is easy to read but hard to analyze. The tidy version has one row per student <em>per month</em>. You will practice this reshaping in this week's assignment, and again with Power Query in week 12.</p>` },
      { heading: 'Faster entry: AutoFill and Freeze Panes', html: `
        <p>In Excel, drag the small square at the bottom right of a selection (the <strong>fill handle</strong>) to continue a pattern: 1, 2, 3…; Jan, Feb, Mar…; Monday, Tuesday… In the practice workbook use <strong>Fill ↓</strong> (<kbd>Ctrl</kbd>+<kbd>D</kbd>) to copy the top cell down a selection.</p>
        <p><strong>View › Freeze Panes</strong> keeps the header row visible while you scroll a long table. (The practice workbook always keeps column letters and row numbers visible.)</p>` },
    ],
    keyTerms: [
      ['Nominal', 'Categories with no natural order.'],
      ['Ordinal', 'Categories with a meaningful order.'],
      ['Discrete', 'Countable whole-number values.'],
      ['Continuous', 'Measured values that can be any number in a range.'],
      ['Structured data', 'Data organized into fixed rows and columns.'],
      ['Number format', 'A display rule; it never changes the stored value.'],
      ['Tidy data', 'One variable per column, one observation per row, one value per cell.'],
    ],
    practice: [
      {
        id: 'p2a', title: 'Workbook 2A · What type is it?',
        intro: 'Type Nominal, Ordinal, Discrete or Continuous in each green cell. The column only accepts those four words: that is data validation, which you will build yourself in week 5.',
        tools: 'basic',
        workbook: { sheets: [{ name: 'Types', rows: typeRows, lock: 'A1:C1', cells: Object.fromEntries(typeRows.slice(1).map((r, i) => ['A' + (i + 2), { input: r[0], locked: true }])), widths: { A: 210, B: 120, C: 330 }, validations: [{ range: 'C2:C9', type: 'list', values: ['Nominal', 'Ordinal', 'Discrete', 'Continuous'] }], styles: { 'C2:C9': { fill: 'answer' } } }] },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Types', range: 'C2:C9', expect: typeAnswers.map((x) => [x]), points: 8, prompt: 'Classify all eight fields in C2:C9.' },
        ],
      },
      {
        id: 'p2b', title: 'Workbook 2B · Fix the types',
        intro: 'This export stored everything as text. Retype the values in the green cells so Excel stores the right type. Watch the alignment flip from left to right when a number becomes a number. Keep the ID column as text.',
        tools: 'basic',
        workbook: {
          sheets: [{
            name: 'Export',
            rows: [
              ['StudentID', 'GPA', 'Enrolled On', 'Lunch Balance', 'Zip Code'],
              ["'00417", "'3.5", "'8/20/2025", "'12.50", '2134'],
              ["'00418", "'3.82", "'8/20/2025", "'4.25", '2139'],
              ["'00419", "'2.9", "'1/6/2026", "'0.75", '1002'],
              ["'00420", "'3.1", "'8/21/2025", "'20", '2215'],
            ],
            lock: 'A1:E1',
            cells: { A2: { input: "'00417", locked: true }, A3: { input: "'00418", locked: true }, A4: { input: "'00419", locked: true }, A5: { input: "'00420", locked: true } },
            styles: { 'B2:E5': { fill: 'answer' } },
            widths: { B: 70, C: 100, D: 105, E: 85 },
          }],
        },
        questions: [
          { id: 'q1', type: 'cellType', sheet: 'Export', range: 'B2:B5', valueType: 'number', expect: [3.5, 3.82, 2.9, 3.1], points: 2, prompt: 'Make GPA (B2:B5) real numbers.', hint: 'Click the cell and retype the value without the apostrophe.' },
          { id: 'q2', type: 'cellType', sheet: 'Export', range: 'C2:C5', valueType: 'date', expect: [45889, 45889, 46028, 45890], points: 2, prompt: 'Make Enrolled On (C2:C5) real dates.' },
          { id: 'q3', type: 'cellType', sheet: 'Export', range: 'D2:D5', valueType: 'number', expect: [12.5, 4.25, 0.75, 20], points: 2, prompt: 'Make Lunch Balance (D2:D5) numbers…' },
          { id: 'q4', type: 'format', sheet: 'Export', range: 'D2:D5', format: 'currency', points: 1, prompt: '…and format them as Currency.' },
          { id: 'q5', type: 'cellType', sheet: 'Export', range: 'E2:E5', valueType: 'text', expect: ['02134', '02139', '01002', '02215'], points: 3, prompt: 'The zip codes lost their leading zero. Retype them as text with the zero (e.g. \'02134).' },
        ],
      },
    ],
    assignment: {
      id: 'hw-2', title: 'Tidy the reading log', points: 20,
      html: `<p>The library sent a "wide" report that is easy to read but hard to analyze. Rebuild it as a <strong>tidy</strong> table on the <code>Tidy</code> sheet: one row per student per month.</p>
        <ol><li>Columns: <code>Student</code>, <code>Month</code>, <code>Books</code> (already in row 1).</li>
        <li>List students in the same order as the report, and months in order Jan, Feb, Mar for each student. That is 12 rows (A2:C13).</li>
        <li>Store Books as numbers. Do not copy the Total row: totals are not data.</li></ol>`,
      workbook: {
        sheets: [
          { name: 'Report', rows: [['Reading Log, Winter', null, null, null], ['Student', 'Jan', 'Feb', 'Mar'], ['Maya', 4, 3, 5], ['Luis', 1, 2, 2], [null, null, null, null], ['Aisha', 6, 5, 7], ['Ethan', 2, 0, 3], ['Total', 13, 10, 17]], lock: true, styles: { A1: { bold: true }, 'A2:D2': { bold: true }, 'A8:D8': { bold: true } } },
          { name: 'Tidy', rows: [['Student', 'Month', 'Books']], lock: 'A1:C1', styles: { 'A2:C13': { fill: 'answer' } } },
        ],
      },
      questions: (() => {
        const students = ['Maya', 'Luis', 'Aisha', 'Ethan'];
        const months = ['Jan', 'Feb', 'Mar'];
        const books = { Maya: [4, 3, 5], Luis: [1, 2, 2], Aisha: [6, 5, 7], Ethan: [2, 0, 3] };
        const sCol = [], mCol = [], bCol = [];
        students.forEach((s) => months.forEach((m, i) => { sCol.push([s]); mCol.push([m]); bCol.push(books[s][i]); }));
        return [
          { id: 'a1', type: 'range', sheet: 'Tidy', range: 'A2:A13', expect: sCol, points: 6, prompt: 'Student column, A2:A13.' },
          { id: 'a2', type: 'range', sheet: 'Tidy', range: 'B2:B13', expect: mCol, points: 6, prompt: 'Month column, B2:B13.' },
          { id: 'a3', type: 'cellType', sheet: 'Tidy', range: 'C2:C13', valueType: 'number', expect: bCol, points: 8, prompt: 'Books column, C2:C13, stored as numbers.' },
        ];
      })(),
    },
    quiz: 'quiz-2',
  });

  // ---------------------------------------------------------------- Week 3
  const signups = [
    ['Name', 'Grade', 'Email', 'Club', 'Hours'],
    ['Maya Patel', 12, 'mpatel@school.org', 'Robotics', 14],
    ['Luis Garcia', 11, null, 'Debate', 6],
    ['Aisha Bello', 12, 'abello@school.org', 'Robotics', 22],
    ['Ethan Nguyen', null, 'enguyen@school.org', 'Art', 9],
    ['Grace Kim', 11, 'gkim@school.org', 'Debate', null],
    ['Noah Smith', 12, null, 'Robotics', 11],
    ['Sofia Rossi', 12, 'srossi@school.org', 'Art', 17],
    ['Kai Tanaka', 11, 'ktanaka@school.org', 'Coding', 25],
    ['Zoe Martin', 12, null, 'Coding', 8],
    ['Omar Haddad', 12, 'ohaddad@school.org', 'Debate', 12],
    ['Lily Chen', 11, 'lchen@school.org', 'Art', 5],
    ['Ben Walker', null, 'bwalker@school.org', 'Coding', 19],
    ['Chloe Davis', 12, 'cdavis@school.org', 'Robotics', 16],
    ['Diego Lopez', 11, null, 'Debate', 7],
    ['Hana Sato', 12, 'hsato@school.org', 'Coding', 21],
    ['Ivy Brooks', 11, 'ibrooks@school.org', 'Art', null],
    ['Jack Ellis', 12, 'jellis@school.org', 'Robotics', 13],
    ['Kira Novak', 11, 'knovak@school.org', 'Coding', 10],
    ['Leo Park', 12, null, 'Debate', 4],
    ['Mia Rivera', 11, 'mrivera@school.org', 'Art', 15],
  ];

  C.lessons.push({
    week: 3, unit: 1, id: 'week-3',
    title: 'Where data comes from and how good it is',
    summary: 'Primary and secondary sources, sampling and bias, the dimensions of data quality, and your first formulas: SUM, AVERAGE, COUNT and friends, with relative and absolute references.',
    objectives: [
      'Distinguish primary from secondary data and name common collection methods.',
      'Identify sampling bias, response bias and survivorship bias in a scenario.',
      'Score a dataset on accuracy, completeness, consistency, timeliness, validity and uniqueness.',
      'Write formulas with operators and the functions SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK and ROUND.',
      'Explain and use relative vs absolute references ($), and fill a formula down.',
    ],
    sections: [
      { heading: 'Where data comes from', html: `
        <p><strong>Primary data</strong> is collected by you for your purpose: a survey you write, a sensor you install, sales your register records. <strong>Secondary data</strong> was collected by someone else for another purpose: census data, a city's open-data portal, a sports statistics site.</p>
        <p>Common collection methods: transactions (every card swipe), forms and surveys, sensors and devices, web and app logs, manual observation, and purchased or public datasets. Each has strengths: transactions are complete but only cover what was bought; surveys ask exactly what you want but people may not answer honestly.</p>` },
      { heading: 'Sampling and bias', html: `
        <p>Usually you cannot measure everyone (the <strong>population</strong>), so you measure some of them (a <strong>sample</strong>). A sample is only useful if it represents the population.</p>
        <ul>
          <li><strong>Selection (sampling) bias:</strong> who you asked is not representative. Surveying only students in the library about study habits.</li>
          <li><strong>Response bias:</strong> people answer inaccurately: they want to look good, or the question is leading ("Don't you agree the cafeteria needs better food?").</li>
          <li><strong>Non-response bias:</strong> the people who skip the survey differ from those who answer.</li>
          <li><strong>Survivorship bias:</strong> you only see the cases that "survived". Studying only successful companies to learn what makes companies successful.</li>
        </ul>
        <div class="callout">Bias is not fixed by collecting <em>more</em> data. A million biased responses are still biased.</div>` },
      { heading: 'The dimensions of data quality', html: `
        <table class="notes-table">
          <thead><tr><th>Dimension</th><th>Question it asks</th><th>Example problem</th></tr></thead>
          <tbody>
            <tr><td><strong>Accuracy</strong></td><td>Does it match reality?</td><td>A student's grade is recorded as 21 instead of 12</td></tr>
            <tr><td><strong>Completeness</strong></td><td>Is anything missing?</td><td>5 of 20 sign-ups have no email</td></tr>
            <tr><td><strong>Consistency</strong></td><td>Is the same thing written the same way?</td><td>"Robotics", "robotics", "Robotics Club"</td></tr>
            <tr><td><strong>Timeliness</strong></td><td>Is it current enough for the decision?</td><td>Using last year's roster to plan this year's buses</td></tr>
            <tr><td><strong>Validity</strong></td><td>Does it follow the rules for that field?</td><td>An email with no @, a grade of 15</td></tr>
            <tr><td><strong>Uniqueness</strong></td><td>Is each record recorded once?</td><td>The same student signed up twice</td></tr>
          </tbody>
        </table>
        <p>Quality can be <em>measured</em>. Completeness of a column = filled cells ÷ total cells. In Workbook 3A you will compute it with a formula.</p>` },
      { heading: 'Formulas', html: `
        <p>A <strong>formula</strong> starts with <code>=</code> and tells Excel to calculate. When the inputs change, the result updates automatically, which is why formulas beat typing answers.</p>
        <table class="notes-table">
          <thead><tr><th>Operator</th><th>Meaning</th><th>Example</th></tr></thead>
          <tbody>
            <tr><td><code>+ - * /</code></td><td>add, subtract, multiply, divide</td><td><code>=B2*C2</code></td></tr>
            <tr><td><code>^</code></td><td>exponent</td><td><code>=2^3</code> → 8</td></tr>
            <tr><td><code>&amp;</code></td><td>join text</td><td><code>=A2&amp;" "&amp;B2</code></td></tr>
            <tr><td><code>= &lt;&gt; &gt; &lt; &gt;= &lt;=</code></td><td>compare (TRUE/FALSE)</td><td><code>=E2&gt;10</code></td></tr>
          </tbody>
        </table>
        <p>Excel follows the order of operations: parentheses, exponents, multiply/divide, add/subtract. <code>=2+3*4</code> is 14; <code>=(2+3)*4</code> is 20.</p>
        <p>While typing a formula, <strong>click a cell</strong> to insert its address, or drag to insert a range. The cells you reference light up in color.</p>` },
      { heading: 'Your first functions', html: `
        <table class="notes-table">
          <thead><tr><th>Function</th><th>Returns</th><th>Example</th></tr></thead>
          <tbody>
            <tr><td><code>SUM(range)</code></td><td>Total of the numbers</td><td><code>=SUM(E2:E21)</code></td></tr>
            <tr><td><code>AVERAGE(range)</code></td><td>Mean of the numbers (blanks skipped)</td><td><code>=AVERAGE(E2:E21)</code></td></tr>
            <tr><td><code>MIN</code> / <code>MAX</code></td><td>Smallest / largest number</td><td><code>=MAX(E2:E21)</code></td></tr>
            <tr><td><code>COUNT(range)</code></td><td>How many cells contain <em>numbers</em></td><td><code>=COUNT(B2:B21)</code></td></tr>
            <tr><td><code>COUNTA(range)</code></td><td>How many cells are <em>not empty</em></td><td><code>=COUNTA(A2:A21)</code></td></tr>
            <tr><td><code>COUNTBLANK(range)</code></td><td>How many cells are empty</td><td><code>=COUNTBLANK(C2:C21)</code></td></tr>
            <tr><td><code>ROUND(number, digits)</code></td><td>Rounds to the given decimals</td><td><code>=ROUND(H5, 1)</code></td></tr>
          </tbody>
        </table>
        <p>COUNT vs COUNTA is a classic quiz question: COUNT counts numbers only; COUNTA counts anything that is not blank.</p>` },
      { heading: 'Relative and absolute references', html: `
        <p>When you copy or fill a formula, Excel shifts its references. <code>=E2/E22</code> filled down becomes <code>=E3/E23</code>, <code>=E4/E24</code>… That is a <strong>relative</strong> reference, and usually what you want.</p>
        <p>Sometimes one reference must stay put, like the total in E22 when you compute each club member's share of all hours. Add dollar signs to make it <strong>absolute</strong>: <code>=E2/$E$22</code>. Filled down, it becomes <code>=E3/$E$22</code>, <code>=E4/$E$22</code>…</p>
        <table class="notes-table">
          <thead><tr><th>Reference</th><th>Column</th><th>Row</th></tr></thead>
          <tbody>
            <tr><td><code>E22</code></td><td>moves</td><td>moves</td></tr>
            <tr><td><code>$E$22</code></td><td>locked</td><td>locked</td></tr>
            <tr><td><code>E$22</code></td><td>moves</td><td>locked</td></tr>
            <tr><td><code>$E22</code></td><td>locked</td><td>moves</td></tr>
          </tbody>
        </table>
        <div class="callout tip">While editing a formula, put the cursor on a reference and press <kbd>F4</kbd> to cycle through the four versions. It works in Excel and in the practice workbook.</div>` },
      { heading: 'Reading error values', html: `
        <ul>
          <li><code>#DIV/0!</code>: dividing by zero or by an empty cell.</li>
          <li><code>#NAME?</code>: a misspelled function name (<code>=SUMM(…)</code>) or text without quotes.</li>
          <li><code>#VALUE!</code>: the wrong type, like adding a number to text.</li>
          <li><code>#REF!</code>: the formula points to a cell that no longer exists.</li>
          <li><code>#N/A</code>: a lookup found no match (week 7).</li>
        </ul>
        <p>Click an error cell in the practice workbook to see a plain-English explanation under the formula bar.</p>` },
    ],
    keyTerms: [
      ['Primary data', 'Collected first-hand for your purpose.'],
      ['Secondary data', 'Collected by someone else for another purpose.'],
      ['Sampling bias', 'The sample does not represent the population.'],
      ['Completeness', 'The share of required values that are present.'],
      ['Formula', 'An instruction beginning with = that calculates a result.'],
      ['Relative reference', 'A reference that shifts when copied (A1).'],
      ['Absolute reference', 'A reference locked with $ that does not shift ($A$1).'],
    ],
    practice: [
      {
        id: 'p3a', title: 'Workbook 3A · Audit the club sign-ups',
        intro: 'Use formulas (start with =) in the green cells. Click cells or drag ranges while typing to insert addresses.',
        tools: 'basic',
        workbook: {
          sheets: [{
            name: 'Signups',
            rows: signups.map((r, i) => (i === 0 ? [...r, null, 'Measure', 'Formula'] : [...r])),
            cells: { G2: 'Number of sign-ups (COUNTA of names)', G3: 'Missing emails (COUNTBLANK)', G4: 'Email completeness (1 − missing ÷ sign-ups)', G5: 'Total volunteer hours', G6: 'Average hours (rounded to 1 decimal)', G7: 'Most hours by one student', G8: 'Students with a recorded grade (COUNT)' },
            formats: { H4: 'percent' },
            lock: 'A1:H1',
            widths: { A: 110, C: 160, G: 300, H: 90 },
            styles: { 'H2:H8': { fill: 'answer' } },
          }],
        },
        questions: [
          { id: 'q1', type: 'cell', sheet: 'Signups', cell: 'H2', solution: '=COUNTA(A2:A21)', functions: ['COUNTA'], prompt: 'H2: count the sign-ups with COUNTA.' },
          { id: 'q2', type: 'cell', sheet: 'Signups', cell: 'H3', solution: '=COUNTBLANK(C2:C21)', functions: ['COUNTBLANK'], prompt: 'H3: count missing emails with COUNTBLANK.' },
          { id: 'q3', type: 'cell', sheet: 'Signups', cell: 'H4', solution: '=1-H3/H2', requireFormula: true, tolerance: 0.001, prompt: 'H4: email completeness as a percent.' },
          { id: 'q4', type: 'cell', sheet: 'Signups', cell: 'H5', solution: '=SUM(E2:E21)', functions: ['SUM'], prompt: 'H5: total hours with SUM.' },
          { id: 'q5', type: 'cell', sheet: 'Signups', cell: 'H6', solution: '=ROUND(AVERAGE(E2:E21),1)', functions: ['AVERAGE', 'ROUND'], prompt: 'H6: average hours, rounded to 1 decimal with ROUND.' },
          { id: 'q6', type: 'cell', sheet: 'Signups', cell: 'H7', solution: '=MAX(E2:E21)', functions: ['MAX'], prompt: 'H7: the most hours with MAX.' },
          { id: 'q7', type: 'cell', sheet: 'Signups', cell: 'H8', solution: '=COUNT(B2:B21)', functions: ['COUNT'], prompt: 'H8: how many students have a grade, with COUNT.' },
        ],
      },
      {
        id: 'p3b', title: 'Workbook 3B · Relative vs absolute references',
        intro: 'E22 holds the total hours. In F2, write a formula for Maya\'s share of all hours, then select F2:F21 and press Fill ↓ (Ctrl+D). If you see #DIV/0!, your reference to E22 moved: make it absolute with F4.',
        tools: 'basic',
        workbook: {
          sheets: [{
            name: 'Share',
            rows: [['Name', 'Grade', 'Email', 'Club', 'Hours', 'Share of all hours'], ...signups.slice(1).map((r) => [r[0], r[1], r[2], r[3], r[4] === null ? 0 : r[4]])],
            cells: { D22: { input: 'Total', locked: true, style: { bold: true } }, E22: { input: '=SUM(E2:E21)', locked: true, style: { bold: true } } },
            formats: { 'F2:F21': 'percent1' },
            lock: 'A1:F1',
            widths: { A: 110, C: 160, F: 130 },
            styles: { 'F2:F21': { fill: 'answer' } },
          }],
        },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Share', range: 'F2:F21', solution: '=E2/$E$22', requireFormula: true, mustReference: ['E$22'], points: 5, prompt: 'F2:F21: each student\'s share of total hours, using an absolute reference to E22.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-3', title: 'Data quality report', points: 25,
      html: `<p>The athletics office exported its equipment check-out log. Before anyone uses it, write a short quality report.</p>
        <ol><li>On the <code>Report</code> sheet, fill each green cell with a <strong>formula</strong> that measures the data on the <code>Checkouts</code> sheet.</li>
        <li>In <code>B11</code>, write two sentences: which quality dimension is the biggest problem, and what you would fix first.</li></ol>`,
      workbook: {
        sheets: [
          {
            name: 'Checkouts',
            rows: [
              ['CheckoutID', 'Student', 'Item', 'Days Out', 'Returned'],
              ['C01', 'Maya Patel', 'Basketball', 3, 'Yes'], ['C02', 'Luis Garcia', 'Soccer Ball', 2, 'Yes'], ['C03', null, 'Tennis Racket', 5, 'No'],
              ['C04', 'Ethan Nguyen', 'basketball', 1, 'Yes'], ['C05', 'Grace Kim', 'Volleyball', null, 'Yes'], ['C06', 'Noah Smith', 'Soccer Ball', 40, 'No'],
              ['C07', 'Sofia Rossi', 'Tennis Racket', 2, null], ['C08', 'Kai Tanaka', 'Basketball', 4, 'Yes'], ['C09', 'Zoe Martin', 'Volleyball', 3, 'Yes'],
              ['C10', null, 'Basketball', 6, 'No'], ['C11', 'Omar Haddad', 'Soccer Ball', 1, 'Yes'], ['C12', 'Lily Chen', 'Volleyball', null, 'Yes'],
              ['C13', 'Ben Walker', 'Tennis Racket', 2, 'Yes'], ['C14', 'Chloe Davis', 'Basketball', 7, null], ['C15', 'Diego Lopez', 'Soccer Ball', 3, 'Yes'],
            ],
            lock: true, widths: { B: 110, C: 110 },
          },
          {
            name: 'Report',
            rows: [
              ['Measure', 'Value'],
              ['Total records', null], ['Missing student names', null], ['Missing Days Out', null], ['Missing Returned', null],
              ['Student name completeness', null], ['Average days out', null], ['Longest checkout (days)', null], ['Shortest checkout (days)', null], [null, null],
              ['What should be fixed first?', null],
            ],
            lock: 'A1:A11', formats: { B6: 'percent' }, widths: { A: 220, B: 420 },
            styles: { 'B2:B9': { fill: 'answer' }, B11: { fill: 'answer' } },
          },
        ],
      },
      questions: [
        { id: 'a1', type: 'cell', sheet: 'Report', cell: 'B2', solution: '=COUNTA(Checkouts!A2:A16)', requireFormula: true, points: 3, prompt: 'Total records.' },
        { id: 'a2', type: 'cell', sheet: 'Report', cell: 'B3', solution: '=COUNTBLANK(Checkouts!B2:B16)', requireFormula: true, points: 3, prompt: 'Missing student names.' },
        { id: 'a3', type: 'cell', sheet: 'Report', cell: 'B4', solution: '=COUNTBLANK(Checkouts!D2:D16)', requireFormula: true, points: 3, prompt: 'Missing Days Out.' },
        { id: 'a4', type: 'cell', sheet: 'Report', cell: 'B5', solution: '=COUNTBLANK(Checkouts!E2:E16)', requireFormula: true, points: 3, prompt: 'Missing Returned.' },
        { id: 'a5', type: 'cell', sheet: 'Report', cell: 'B6', solution: '=1-B3/B2', requireFormula: true, tolerance: 0.001, points: 3, prompt: 'Student name completeness (percent).' },
        { id: 'a6', type: 'cell', sheet: 'Report', cell: 'B7', solution: '=AVERAGE(Checkouts!D2:D16)', requireFormula: true, tolerance: 0.01, points: 2, prompt: 'Average days out.' },
        { id: 'a7', type: 'cell', sheet: 'Report', cell: 'B8', solution: '=MAX(Checkouts!D2:D16)', requireFormula: true, points: 2, prompt: 'Longest checkout.' },
        { id: 'a8', type: 'cell', sheet: 'Report', cell: 'B9', solution: '=MIN(Checkouts!D2:D16)', requireFormula: true, points: 2, prompt: 'Shortest checkout.' },
        { id: 'a9', type: 'filled', sheet: 'Report', range: 'B11', min: 1, points: 4, prompt: 'Two-sentence recommendation in B11.' },
      ],
      rubric: [['Recommendation names a quality dimension and a specific fix (e.g. the 40-day outlier, "basketball" vs "Basketball")', 'Teacher check']],
    },
    quiz: 'quiz-3',
  });

  // ---------------------------------------------------------------- Week 4
  C.lessons.push({
    week: 4, unit: 1, id: 'week-4',
    title: 'CSV vs Excel, and delimiting data',
    summary: 'What a file format is, what CSV keeps and loses compared with .xlsx, delimiters and text qualifiers, importing without damaging data, and Text to Columns.',
    objectives: [
      'Explain the difference between a .csv file and an .xlsx workbook, and pick the right one for a task.',
      'Identify a delimiter and a text qualifier in raw text.',
      'Import a delimited file while protecting IDs and zip codes as text.',
      'Split a column into several with Text to Columns.',
      'List what is lost when a workbook is saved as CSV.',
    ],
    sections: [
      { heading: 'Two very different files', html: `
        <p>An <strong>.xlsx</strong> file is a workbook: a zipped bundle of XML files that stores multiple sheets, formulas, number formats, colors, charts, tables and pivot tables. Only spreadsheet programs can open it fully.</p>
        <p>A <strong>.csv</strong> (comma-separated values) file is plain text. Each line is one row; commas separate the columns. That is all. Any program can read it: Notepad, Python, a database, Power BI, a website.</p>
        <table class="notes-table">
          <thead><tr><th></th><th>.xlsx workbook</th><th>.csv text file</th></tr></thead>
          <tbody>
            <tr><td>Sheets</td><td>Many</td><td>One table only</td></tr>
            <tr><td>Formulas</td><td>Kept</td><td>Only the results are saved</td></tr>
            <tr><td>Formatting, colors, charts</td><td>Kept</td><td>Lost</td></tr>
            <tr><td>Data types</td><td>Stored with each cell</td><td>Everything is text; the program reading it guesses</td></tr>
            <tr><td>Opens in</td><td>Spreadsheet apps</td><td>Anything that reads text</td></tr>
            <tr><td>Best for</td><td>Analysis, reports, sharing finished work with people</td><td>Moving data between systems, exports, imports, very large simple tables</td></tr>
          </tbody>
        </table>` },
      { heading: 'Delimiters and text qualifiers', html: `
        <p>A <strong>delimiter</strong> is the character that separates fields. Commas are most common, but you will meet <strong>tabs</strong> (.tsv / .txt), <strong>pipes</strong> <code>|</code> (popular because names and addresses rarely contain them) and <strong>semicolons</strong> (common in Europe, where the comma is the decimal point).</p>
        <pre class="raw">StudentID,Name,Zip,Phone
00417,"Patel, Maya",02134,555-0141</pre>
        <p>The name contains a comma, so it is wrapped in double quotes: the <strong>text qualifier</strong>. Without the quotes, a program would split "Patel" and "Maya" into two columns and every column after it would shift.</p>` },
      { heading: 'What goes wrong when you open a CSV carelessly', html: `
        <ul>
          <li><strong>Leading zeros vanish:</strong> <code>02134</code> becomes the number 2134. Student IDs, zip codes and product codes get damaged.</li>
          <li><strong>Long IDs turn into scientific notation:</strong> a 16-digit card or tracking number shows as <code>1.23457E+15</code> and loses digits.</li>
          <li><strong>Dates get reinterpreted:</strong> <code>03/04/2026</code> is March 4 in the US and April 3 in much of the world.</li>
          <li><strong>Text that looks like a formula or number changes:</strong> gene names like MARCH1 famously turned into dates.</li>
        </ul>
        <p>The fix is to <strong>import</strong> instead of double-clicking: in Excel use <strong>Data › From Text/CSV</strong>, check the preview, and set ID-like columns to <strong>Text</strong> before loading. That import window is actually Power Query, which you will use for real in week 11.</p>` },
      { heading: 'Text to Columns', html: `
        <p>When data is already in Excel but crammed into one column ("Patel, Maya"), select the column and use <strong>Data › Text to Columns</strong>. Choose the delimiter, preview the split, and mark any column that must stay text. The pieces land in the columns to the right, so make sure those are empty.</p>` },
      { heading: 'Choosing the format', html: `
        <ul>
          <li>Sending finished analysis to a teacher or manager: <strong>.xlsx</strong> (or PDF).</li>
          <li>Uploading a roster into the school's grade system: <strong>.csv</strong>, the format the system asks for.</li>
          <li>Downloading transactions from a bank or website: usually <strong>.csv</strong>; import it into a workbook to analyze.</li>
          <li>Keeping data that a program will update every month: <strong>.csv</strong> as the raw source, plus a workbook that imports it.</li>
        </ul>
        <div class="callout tip"><strong>Try it:</strong> in Workbook 4A, click <em>Save as CSV</em> and open the downloaded file in Notepad. Your formulas are gone and only the results remain.</div>` },
    ],
    keyTerms: [
      ['CSV', 'Comma-separated values: a plain-text file with one record per line.'],
      ['Delimiter', 'The character that separates fields (comma, tab, pipe, semicolon).'],
      ['Text qualifier', 'Quotes around a field so delimiters inside it are not treated as separators.'],
      ['Import', 'Bringing external data into a workbook with control over types.'],
      ['Text to Columns', 'Splits one column into several using a delimiter.'],
    ],
    practice: [
      {
        id: 'p4a', title: 'Workbook 4A · Import a pipe-delimited roster',
        intro: 'Click From Text/CSV, choose roster_pipe.txt, look at the raw text, and load it. Before loading, tick StudentID and Zip as text so their leading zeros survive. Then try Save as CSV and look at what you get.',
        tools: ['format', 'import', 'csv', 'sheets', 'download', 'reset'],
        workbook: {
          sheets: [{ name: 'Start Here', rows: [['Use Data › From Text/CSV to load roster_pipe.txt into a new sheet.'], ['Keep StudentID and Zip as text.']], headerRow: false, lock: true, widths: { A: 520 } }],
          sources: { 'roster_pipe.txt': D.files['roster_pipe.txt'], 'roster_comma.csv': D.files['roster_comma.csv'], 'roster_tab.txt': D.files['roster_tab.txt'] },
        },
        model: (wb) => wb.importDelimited(wb.sources['roster_pipe.txt'], { delimiter: '|', textColumns: [0, 2, 3], sheetName: 'roster_pipe' }),
        questions: [
          { id: 'q1', type: 'cellType', sheet: 'roster_pipe', range: 'A2:A5', valueType: 'text', expect: ['00417', '00418', '00419', '00420'], points: 2, prompt: 'Sheet roster_pipe: StudentIDs kept as text (00417…).' },
          { id: 'q2', type: 'cellType', sheet: 'roster_pipe', range: 'C2:C5', valueType: 'text', expect: ['02134', '02139', '01002', '02215'], points: 2, prompt: 'Sheet roster_pipe: zip codes kept as text (02134…).' },
          { id: 'q3', type: 'choice', prompt: 'You save a workbook with three sheets and formulas as CSV. What does the CSV contain?', options: ['All three sheets with formulas', 'Only the active sheet, with values instead of formulas', 'Only the formulas, without values', 'An exact copy of the workbook'], answer: 1 },
        ],
      },
      {
        id: 'p4b', title: 'Workbook 4B · Text to Columns',
        intro: 'Column A has "Last, First" names. Select A2:A7, open Text to Columns, choose Comma, and finish. The last names land in A, first names in B.',
        tools: ['format', 'textToColumns', 'download', 'reset'],
        model: (wb) => wb.textToColumns('Names', 'A2:A7', ','),
        workbook: { sheets: [{ name: 'Names', rows: [['Name (Last, First)', 'First'], ['Patel, Maya'], ['Garcia, Luis'], ['Bello, Aisha'], ['Nguyen, Ethan'], ['Kim, Grace'], ['Smith, Noah']], lock: 'A1:B1', widths: { A: 170, B: 120 } }] },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Names', range: 'A2:A7', expect: [['Patel'], ['Garcia'], ['Bello'], ['Nguyen'], ['Kim'], ['Smith']], points: 3, prompt: 'A2:A7 holds last names.' },
          { id: 'q2', type: 'range', sheet: 'Names', range: 'B2:B7', expect: [['Maya'], ['Luis'], ['Aisha'], ['Ethan'], ['Grace'], ['Noah']], points: 3, prompt: 'B2:B7 holds first names (no leading space).' },
        ],
      },
    ],
    assignment: {
      id: 'hw-4', title: 'From delimited file to clean workbook', points: 20,
      html: `<p>The district sent the new roster as a <strong>tab-delimited</strong> text file. Turn it into a clean workbook and explain your choice of format.</p>
        <ol><li>Import <code>roster_tab.txt</code> with From Text/CSV, keeping StudentID and Zip as text.</li>
        <li>Rename the new sheet to <code>Roster</code> (double-click the sheet tab).</li>
        <li>On the <code>Memo</code> sheet, answer the three questions in B2:B4 in one or two sentences each.</li></ol>`,
      tools: ['format', 'import', 'csv', 'sheets', 'download', 'reset'],
      workbook: {
        sheets: [{ name: 'Memo', rows: [['Question', 'Your answer'], ['When would you send a coworker this data as .xlsx?'], ['When would you send it as .csv?'], ['What would have gone wrong if you had double-clicked the file instead of importing it?']], lock: 'A1:A4', widths: { A: 420, B: 420 }, styles: { 'B2:B4': { fill: 'answer' } } }],
        sources: { 'roster_tab.txt': D.files['roster_tab.txt'] },
      },
      model: (wb) => wb.importDelimited(wb.sources['roster_tab.txt'], { delimiter: 'tab', textColumns: [0, 2, 3], sheetName: 'Roster' }),
      questions: [
        { id: 'a1', type: 'cellType', sheet: 'Roster', range: 'A2:A5', valueType: 'text', expect: ['00417', '00418', '00419', '00420'], points: 5, prompt: 'Roster sheet: StudentID kept as text.' },
        { id: 'a2', type: 'cellType', sheet: 'Roster', range: 'C2:C5', valueType: 'text', expect: ['02134', '02139', '01002', '02215'], points: 5, prompt: 'Roster sheet: Zip kept as text.' },
        { id: 'a3', type: 'range', sheet: 'Roster', range: 'B2:B5', expect: [['Maya Patel'], ['Luis Garcia'], ['Aisha Bello'], ['Ethan Nguyen']], points: 4, prompt: 'Roster sheet: names in column B.' },
        { id: 'a4', type: 'filled', sheet: 'Memo', range: 'B2:B4', min: 3, points: 6, prompt: 'All three memo answers written.' },
      ],
    },
    quiz: 'quiz-4',
  });

  // ---------------------------------------------------------------- Week 5
  const books = [
    ['BookID', 'Title', 'Author', 'Genre', 'Copies', 'Status', 'Days Overdue'],
    ['B101', 'The Giver', 'Lois Lowry', 'Fiction', 4, 'Available', 0],
    ['B102', 'Hatchet', 'Gary Paulsen', 'Adventure', 3, 'Checked Out', 6],
    ['B103', 'Educated', 'Tara Westover', 'Memoir', 2, 'Checked Out', 21],
    ['B104', 'Dune', 'Frank Herbert', 'Sci-Fi', 5, 'Available', 0],
    ['B105', 'Wonder', 'R. J. Palacio', 'Fiction', 3, 'Checked Out', 15],
    ['B106', 'Holes', 'Louis Sachar', 'Fiction', 6, 'Available', 0],
    ['B107', 'The Martian', 'Andy Weir', 'Sci-Fi', 2, 'Checked Out', 2],
    ['B108', 'Unbroken', 'Laura Hillenbrand', 'Memoir', 1, 'Lost', 60],
    ['B109', 'Animal Farm', 'George Orwell', 'Fiction', 4, 'Available', 0],
    ['B110', 'Into the Wild', 'Jon Krakauer', 'Adventure', 2, 'Checked Out', 30],
    ['B111', 'Ender\'s Game', 'Orson Scott Card', 'Sci-Fi', 3, 'Available', 0],
    ['B112', 'Night', 'Elie Wiesel', 'Memoir', 4, 'Checked Out', 9],
  ];

  C.lessons.push({
    week: 5, unit: 1, id: 'week-5',
    title: 'Organizing data for others, and Unit 1 review',
    summary: 'Sorting, filtering, Excel Tables and structured references, data validation and conditional formatting. Then a review of Unit 1 before Friday\'s exam.',
    objectives: [
      'Sort a table by one or more columns without breaking rows apart.',
      'Filter rows to answer a question, and explain why filtering does not delete data.',
      'Convert a range to an Excel Table and write formulas with structured references.',
      'Prevent bad entries with data validation.',
      'Draw attention to important values with conditional formatting.',
    ],
    sections: [
      { heading: 'Sorting', html: `
        <p>Sorting reorders <strong>whole rows</strong> by the values in one column: A→Z or smallest→largest (ascending), or the reverse (descending). Always sort the entire table, never one selected column. Sorting a single column scrambles your records, because one student's grade ends up next to another student's name.</p>
        <p>A <strong>multi-level sort</strong> sorts by one column, then breaks ties with another: Genre A→Z, then Title A→Z within each genre. Tell Excel your data has headers so the header row stays on top.</p>` },
      { heading: 'Filtering', html: `
        <p>A filter <strong>hides</strong> rows that do not match, so you can focus on part of the data: only "Checked Out" books, only sales over $20. The hidden rows are still there. Row numbers skip and turn blue in Excel as a reminder, and clearing the filter brings everything back.</p>
        <div class="callout">Filtered data and the status bar: select a filtered column and the status bar sums only the visible rows. Formulas like <code>SUM</code> still include hidden rows.</div>` },
      { heading: 'Excel Tables', html: `
        <p>Select any cell in your data and press <kbd>Ctrl</kbd>+<kbd>T</kbd> (<strong>Format as Table</strong>). A Table:</p>
        <ul>
          <li>gets a name (rename it from Table1 to something meaningful like <code>Books</code>),</li>
          <li>adds banded rows and filter buttons,</li>
          <li><strong>grows automatically</strong> when you type in the row underneath, so formulas and charts include new data,</li>
          <li>lets you write <strong>structured references</strong> by column name instead of cell addresses.</li>
        </ul>
        <table class="notes-table">
          <thead><tr><th>Structured reference</th><th>Means</th></tr></thead>
          <tbody>
            <tr><td><code>=SUM(Books[Copies])</code></td><td>Sum every value in the Copies column of the Books table</td></tr>
            <tr><td><code>=[@Copies]*2</code></td><td>The Copies value in <em>this row</em> (used inside the table)</td></tr>
            <tr><td><code>=COUNTA(Books[Title])</code></td><td>How many books are in the table, even after you add more</td></tr>
          </tbody>
        </table>
        <p>Tables are also how Power Query and Power BI like to receive data, so this habit pays off in Unit 3.</p>` },
      { heading: 'Data validation: quality at the point of entry', html: `
        <p>Week 3 measured quality <em>after</em> the fact. <strong>Data validation</strong> prevents problems <em>before</em> they happen by limiting what can be typed in a cell:</p>
        <ul>
          <li><strong>List</strong>: only Available, Checked Out or Lost (shows a drop-down in Excel).</li>
          <li><strong>Whole number between</strong> 9 and 12 for a Grade column.</li>
          <li><strong>Decimal</strong> between 0 and 4 for GPA; <strong>text length</strong> of exactly 5 for a zip code.</li>
        </ul>
        <p>Each rule directly protects a quality dimension: lists protect <em>consistency</em>, ranges protect <em>validity</em>.</p>` },
      { heading: 'Conditional formatting', html: `
        <p><strong>Conditional formatting</strong> changes a cell's look when its value meets a rule, like highlighting books more than 14 days overdue, marking duplicates, or shading the top 3 sellers. Readers find what matters without reading every number. Use it sparingly: if everything is highlighted, nothing stands out.</p>` },
      { heading: 'Unit 1 review checklist', html: `
        <p>Exam 1 is Friday and covers weeks 1–5 only. You should be able to:</p>
        <ul class="checklist-review">
          <li>Explain data vs information vs knowledge with an example, and the data-to-decision cycle.</li>
          <li>Classify fields as nominal, ordinal, discrete or continuous; spot "numbers" that should be text.</li>
          <li>Explain that dates are stored as numbers and formats only change display.</li>
          <li>Name the six data quality dimensions and match each to an example.</li>
          <li>Write formulas with SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK and ROUND; use $ references and fill down.</li>
          <li>Compare CSV and .xlsx; explain delimiters and text qualifiers; import without losing leading zeros; use Text to Columns.</li>
          <li>Sort, filter, create a Table, write a structured reference, and add validation and conditional formatting.</li>
        </ul>
        <p>Retake any practice workbook from weeks 1–4 with <em>Reset</em>, then try the exam review on the Assessments page.</p>` },
    ],
    keyTerms: [
      ['Sort', 'Reorder whole rows by the values in a column.'],
      ['Filter', 'Temporarily hide rows that do not meet a condition.'],
      ['Excel Table', 'A named, self-expanding range with headers (Ctrl+T).'],
      ['Structured reference', 'A table-and-column reference such as Books[Copies].'],
      ['Data validation', 'Rules that limit what can be entered in a cell.'],
      ['Conditional formatting', 'Formatting applied automatically when a rule is met.'],
    ],
    practice: [
      {
        id: 'p5a', title: 'Workbook 5A · Organize the library inventory',
        intro: 'Work through the tasks in order. Every tool is on the ribbon: Sort, Format as Table, Data Validation and Highlight Rules.',
        tools: 'organize',
        workbook: { sheets: [{ name: 'Inventory', rows: books, widths: { B: 140, C: 140, F: 100, G: 105 }, cells: { I1: { input: 'Total copies', style: { bold: true }, locked: true }, I3: { input: 'Books checked out', style: { bold: true }, locked: true } }, styles: { I2: { fill: 'answer' }, I4: { fill: 'answer' } } }] },
        questions: [
          { id: 'q1', type: 'sorted', sheet: 'Inventory', range: 'A1:G13', column: 'B', order: 'asc', points: 2, prompt: 'Sort the inventory by Title, A→Z (keep whole rows together).' },
          { id: 'q2', type: 'table', sheet: 'Inventory', range: 'A1:G13', name: 'Books', solution: 'Books', points: 2, prompt: 'Format A1:G13 as a Table named Books.' },
          { id: 'q3', type: 'cell', sheet: 'Inventory', cell: 'I2', solution: '=SUM(Books[Copies])', functions: ['SUM'], mustReference: ['Books[Copies]'], points: 2, prompt: 'In I2, total the copies with a structured reference: =SUM(Books[Copies]).' },
          { id: 'q4', type: 'cell', sheet: 'Inventory', cell: 'I4', solution: '=COUNTIF(Books[Status],"Checked Out")', requireFormula: true, points: 2, prompt: 'In I4, count books with Status "Checked Out" (hint: COUNTIF, which you will master next week).' },
          { id: 'q5', type: 'validation', sheet: 'Inventory', range: 'F2:F13', solution: { type: 'list', values: ['Available', 'Checked Out', 'Lost'] }, points: 2, prompt: 'Add a List validation to F2:F13 allowing Available, Checked Out, Lost.' },
          { id: 'q6', type: 'condFormat', sheet: 'Inventory', range: 'G2:G13', solution: { op: '>', value: 14, style: 'red' }, points: 2, prompt: 'Highlight Days Overdue greater than 14 in G2:G13.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-5', title: 'Unit 1 mini-project: clean and summarize volunteer hours', points: 30,
      html: `<p>The service-learning coordinator needs a clean, organized log and a summary. Use everything from Unit 1.</p>
        <ol><li><strong>Remove duplicate rows</strong> (the export repeated some entries).</li>
        <li><strong>Sort</strong> the log by Student, A→Z.</li>
        <li><strong>Format as Table</strong> named <code>Hours</code>.</li>
        <li>Add <strong>validation</strong> to the Type column allowing only Tutoring, Food Bank, Cleanup.</li>
        <li>Fill the summary cells (K2:K5) with formulas.</li></ol>
        <p>The checker grades the summary values, so they must reflect the data <em>after</em> duplicates are removed.</p>`,
      tools: 'organize',
      workbook: (() => {
        const log = [
          ['Student', 'Date', 'Type', 'Hours'],
          ['Zoe Martin', '2/3/2026', 'Tutoring', 2], ['Maya Patel', '2/3/2026', 'Food Bank', 3], ['Luis Garcia', '2/4/2026', 'Cleanup', 1.5],
          ['Maya Patel', '2/3/2026', 'Food Bank', 3], ['Ethan Nguyen', '2/5/2026', 'Tutoring', 2], ['Aisha Bello', '2/6/2026', 'Cleanup', 4],
          ['Zoe Martin', '2/10/2026', 'Food Bank', 3], ['Luis Garcia', '2/4/2026', 'Cleanup', 1.5], ['Grace Kim', '2/11/2026', 'Tutoring', 1],
          ['Aisha Bello', '2/12/2026', 'Food Bank', 2.5], ['Ethan Nguyen', '2/13/2026', 'Cleanup', 2], ['Grace Kim', '2/11/2026', 'Tutoring', 1],
          ['Noah Smith', '2/17/2026', 'Tutoring', 3],
        ];
        return { sheets: [{ name: 'Log', rows: log, widths: { A: 120, B: 90, J: 230 }, cells: { J1: { input: 'Summary', style: { bold: true }, locked: true }, J2: { input: 'Entries (after removing duplicates)', locked: true }, J3: { input: 'Total hours', locked: true }, J4: { input: 'Average hours per entry', locked: true }, J5: { input: 'Tutoring entries', locked: true } }, styles: { 'K2:K5': { fill: 'answer' } } }] };
      })(),
      model: (wb) => { wb.removeDuplicates('Log', 'A1:D14'); wb.sortRange('Log', 'A1:D11', 0, true, true); },
      questions: (() => {
        // computed from the de-duplicated log
        const uniq = [['Zoe Martin', 2], ['Maya Patel', 3], ['Luis Garcia', 1.5], ['Ethan Nguyen', 2], ['Aisha Bello', 4], ['Zoe Martin', 3], ['Grace Kim', 1], ['Aisha Bello', 2.5], ['Ethan Nguyen', 2], ['Noah Smith', 3]];
        const total = uniq.reduce((a, b) => a + b[1], 0);
        return [
          { id: 'a1', type: 'cell', sheet: 'Log', cell: 'K2', expect: uniq.length, solution: '=COUNTA(A2:A11)', requireFormula: true, points: 4, prompt: 'Number of entries after removing duplicates.' },
          { id: 'a2', type: 'cell', sheet: 'Log', cell: 'K3', expect: total, solution: '=SUM(D2:D11)', requireFormula: true, points: 4, prompt: 'Total hours.' },
          { id: 'a3', type: 'cell', sheet: 'Log', cell: 'K4', expect: total / uniq.length, solution: '=AVERAGE(D2:D11)', tolerance: 0.01, requireFormula: true, points: 4, prompt: 'Average hours per entry.' },
          { id: 'a4', type: 'cell', sheet: 'Log', cell: 'K5', expect: 4, solution: '=COUNTIF(C2:C11,"Tutoring")', requireFormula: true, points: 4, prompt: 'Number of Tutoring entries.' },
          { id: 'a5', type: 'table', sheet: 'Log', range: 'A1:D11', name: 'Hours', points: 5, prompt: 'Table named Hours covering the cleaned data.' },
          { id: 'a6', type: 'validation', sheet: 'Log', range: 'C2:C11', solution: { type: 'list', values: ['Tutoring', 'Food Bank', 'Cleanup'] }, points: 4, prompt: 'List validation on Type.' },
          { id: 'a7', type: 'cell', sheet: 'Log', cell: 'A2', expect: 'Aisha Bello', points: 5, prompt: 'Log sorted by Student A→Z (Aisha Bello first).' },
        ];
      })(),
    },
    quiz: 'quiz-5',
    exam: 'exam-1',
  });

  if (isNode) module.exports = C;
})(typeof globalThis !== 'undefined' ? globalThis : this);
