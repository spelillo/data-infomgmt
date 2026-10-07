/*
 * unit3.js — Unit 3: Repeatable data work with Power Query and Power BI (weeks 11–15).
 */
(function (root) {
  'use strict';
  const isNode = typeof module === 'object' && module.exports;
  const C = isNode ? require('./course.js') : root.DIM.Content;
  const D = C.data;

  const monthFiles = ['sales_jan.csv', 'sales_feb.csv', 'sales_mar.csv'];
  const pick = (...names) => Object.fromEntries(names.map((n) => [n, D.files[n]]));

  // The cleaning recipe students build in week 11 and reuse after that.
  const cleanSteps = () => [
    { type: 'removeBlankRows' },
    { type: 'removeDuplicates', columns: [] },
    { type: 'transform', column: 'Store', op: 'trim' },
    { type: 'transform', column: 'Store', op: 'proper' },
    { type: 'changeType', columns: { 'Order ID': 'whole', 'Order Date': 'date', Qty: 'whole', 'Unit Price': 'currency' } },
    { type: 'addColumn', name: 'Revenue', formula: '[Qty]*[Unit Price]', dataType: 'currency' },
  ];
  const startSheet = (lines) => ({ name: 'Start Here', rows: lines.map((l) => [l]), headerRow: false, lock: true, widths: { A: 640 } });

  // ---------------------------------------------------------------- Week 11
  C.lessons.push({
    week: 11, unit: 3, id: 'week-11',
    title: 'Power Query basics',
    summary: 'Why repeatable cleaning matters, connecting to a file, the Power Query Editor, applied steps, data types, and refreshing a query when next month\'s file arrives.',
    objectives: [
      'Explain what Power Query is and why a recorded process beats cleaning by hand.',
      'Connect to a CSV file and open the Power Query Editor.',
      'Apply, inspect, reorder and delete steps in the Applied Steps list.',
      'Set column data types and explain why types matter.',
      'Point a query at a new file and refresh to repeat every step.',
    ],
    sections: [
      { heading: 'Remember week 8?', html: `
        <p>You cleaned February's sign-ups by hand, then March's file arrived and you did it all again. Power Query fixes this. It is a tool built into Excel (and Power BI) that <strong>records every cleaning step</strong> as you do it. When new data arrives, you press <strong>Refresh</strong> and every step runs again automatically.</p>
        <p>Power Query follows a pattern called <strong>ETL</strong>: <em>Extract</em> data from a source, <em>Transform</em> it (clean, reshape, combine), and <em>Load</em> it somewhere useful (a sheet, a PivotTable, a Power BI model).</p>` },
      { heading: 'Getting data', html: `
        <p>In Excel: <strong>Data › Get Data › From File › From Text/CSV</strong>, pick the file, then click <strong>Transform Data</strong> (not Load) to open the Power Query Editor. In the practice workbook, click <strong>Power Query</strong> on the ribbon, then <strong>+ New query</strong>.</p>
        <p>Sources can be a file, a <strong>folder</strong> of files (week 12), a table or sheet in the workbook, a database, or a web page. The source is always the first entry in Applied Steps.</p>` },
      { heading: 'The Power Query Editor', html: `
        <ul>
          <li><strong>Preview grid</strong> in the middle: the data after the selected step. Click a column header to select it.</li>
          <li><strong>Ribbon</strong> on top: the transformations (Remove Rows, Change Type, Replace Values, Format, Split Column, Add Column, Group By…).</li>
          <li><strong>Applied Steps</strong> on the right: every transformation in order. Click a step to see the data at that point. Delete a step with ×. Steps run top to bottom, so order matters: you cannot trim a column after you deleted it.</li>
          <li><strong>Close &amp; Load</strong>: runs the query and writes the result to a sheet.</li>
        </ul>
        <div class="callout">Power Query <strong>never changes the source file</strong>. The steps are applied to a copy each time. That makes the process safe, repeatable and auditable: anyone can open Applied Steps and see exactly what was done to the data.</div>` },
      { heading: 'Data types', html: `
        <p>Each column has a type shown by an icon in the header: ABC for text, 123 for whole numbers, 1.2 for decimals, $ for currency, a calendar for dates. Raw CSV columns arrive as text. Use <strong>Transform › Data Type</strong> (or the "Detect data types" option) so numbers can be added and dates can be grouped. Keep IDs and zip codes as text.</p>
        <p>If a value cannot be converted (like "ten" in a number column), the cell becomes <code>Error</code>. Errors are information: find out why before using <strong>Remove Errors</strong>.</p>` },
      { heading: 'A typical cleaning recipe', html: `
        <ol>
          <li><strong>Remove Blank Rows</strong> left by the export.</li>
          <li><strong>Remove Duplicates</strong> (select no column to compare entire rows).</li>
          <li><strong>Format › Trim</strong> then <strong>Format › Capitalize Each Word</strong> on text columns like Store.</li>
          <li><strong>Change Type</strong>: dates to Date, money to Currency, counts to Whole Number.</li>
          <li><strong>Add Custom Column</strong>: <code>Revenue = [Qty] * [Unit Price]</code>. Column names go in square brackets.</li>
          <li><strong>Close &amp; Load</strong>.</li>
        </ol>
        <p>Compare this to week 8: the same ideas as TRIM, PROPER and Remove Duplicates, but recorded once.</p>` },
      { heading: 'Refresh: the payoff', html: `
        <p>Next month, put the new file where the old one was (or change the source: in Excel, <strong>Data source settings</strong>; in the practice editor, the ⚙ next to Source) and click <strong>Refresh All</strong>. Every step runs on the new data. A process that took 20 minutes by hand now takes 2 seconds, and it is done the same way every time.</p>` },
    ],
    keyTerms: [
      ['Power Query', 'Excel and Power BI\'s tool for recording repeatable data transformations.'],
      ['ETL', 'Extract, Transform, Load.'],
      ['Applied Steps', 'The ordered list of transformations a query performs.'],
      ['Data type', 'The kind of value a column holds (text, whole number, decimal, date, currency).'],
      ['Refresh', 'Re-run a query on the current source data.'],
      ['Custom column', 'A new column calculated from other columns, e.g. [Qty]*[Unit Price].'],
    ],
    practice: [
      {
        id: 'p11a', title: 'Workbook 11A · Your first query',
        intro: 'Open Power Query, create a New query From Text/CSV on sales_jan.csv and name it Jan. Apply the cleaning recipe from the notes, then Close & Load.',
        tools: ['query', 'sheets', 'download', 'reset'],
        workbook: { sheets: [startSheet(['Click Power Query on the ribbon to begin.', 'Recipe: Remove Blank Rows · Remove Duplicates · Trim and Capitalize Store · Change Types · Add Revenue = [Qty]*[Unit Price] · Close & Load'])], sources: pick('sales_jan.csv', 'sales_feb.csv') },
        questions: [
          { id: 'q1', type: 'query', solution: { name: 'Jan', source: { kind: 'file', file: 'sales_jan.csv' }, steps: cleanSteps() }, steps: ['removeBlankRows', 'removeDuplicates', 'transform', 'changeType', 'addColumn'], loaded: true, points: 6, prompt: 'Query on sales_jan.csv with all five kinds of steps, loaded to a sheet.' },
          { id: 'q2', type: 'cell', sheet: 'Jan', cell: 'A1', expect: 'Order ID', solution: null, points: 1, prompt: 'The query is loaded to a sheet named Jan.' },
          { id: 'q3', type: 'choice', prompt: 'You delete the "Removed Duplicates" step. What happens to the CSV file on disk?', options: ['The duplicates are put back in the file', 'Nothing: Power Query never changes the source', 'The file is deleted', 'The file is re-sorted'], answer: 1 },
        ],
      },
      {
        id: 'p11b', title: 'Workbook 11B · February arrives: refresh',
        intro: 'This workbook already has the cleaned "Sales" query, loaded from January. February\'s file just arrived. Open Power Query, click ⚙ next to Source, choose sales_feb.csv, and Close & Load. Every step re-runs on the new file.',
        tools: ['query', 'sheets', 'download', 'reset'],
        workbook: { sheets: [startSheet(['The Sales sheet is loaded from sales_jan.csv. Point the query at sales_feb.csv and reload.'])], sources: pick('sales_jan.csv', 'sales_feb.csv'), queries: [{ id: 'q-Sales', name: 'Sales', source: { kind: 'file', file: 'sales_jan.csv' }, steps: cleanSteps() }] },
        model: (wb) => { wb.loadQuery('q-Sales'); },
        questions: [
          { id: 'q1', type: 'query', solution: { name: 'Sales', source: { kind: 'file', file: 'sales_feb.csv' }, steps: cleanSteps() }, requireSource: true, loaded: true, points: 4, prompt: 'The Sales query points at sales_feb.csv and is reloaded.' },
          { id: 'q2', type: 'number', prompt: 'After the refresh, how many orders does the Sales sheet show?', answer: 16 },
        ],
      },
    ],
    assignment: {
      id: 'hw-11', title: 'Store totals query for March', points: 25,
      html: `<p>Build a query that cleans March's export and summarizes it by store, so the manager can refresh it every month.</p>
        <ol><li>New query from <code>sales_mar.csv</code>, named <code>StoreTotals</code>.</li>
        <li>Clean it: remove blank rows and duplicates, trim and capitalize Store, set types, add Revenue.</li>
        <li><strong>Group By</strong> Store with the Sum of Revenue as a column named <code>Revenue</code>. Close &amp; Load.</li>
        <li>On <code>Notes</code>, type the Cafeteria total in B2, then explain in B3 why the Store column needed trimming <em>before</em> grouping.</li></ol>`,
      tools: ['query', 'format', 'sheets', 'download', 'reset'],
      workbook: {
        sheets: [{ name: 'Notes', rows: [['Question', 'Answer'], ['Cafeteria revenue in March'], ['Why trim and capitalize before Group By?']], formats: { B2: 'currency' }, lock: 'A1:A3', widths: { A: 300, B: 460 }, styles: { 'B2:B3': { fill: 'answer' } } }],
        sources: pick('sales_mar.csv'),
      },
      questions: [
        { id: 'a1', type: 'query', solution: { name: 'StoreTotals', source: { kind: 'file', file: 'sales_mar.csv' }, steps: [...cleanSteps(), { type: 'groupBy', by: ['Store'], aggregations: [{ name: 'Revenue', op: 'sum', column: 'Revenue' }] }] }, steps: ['removeBlankRows', 'removeDuplicates', 'transform', 'changeType', 'addColumn', 'groupBy'], loaded: true, points: 14, prompt: 'StoreTotals query: cleaning steps plus Group By, loaded.' },
        { id: 'a2', type: 'cell', sheet: 'Notes', cell: 'B2', solution: '=SUMIF(StoreTotals!A:A,"Cafeteria",StoreTotals!B:B)', tolerance: 0.005, points: 5, prompt: 'Cafeteria revenue in March.' },
        { id: 'a3', type: 'filled', sheet: 'Notes', range: 'B3', min: 1, points: 6, prompt: 'Explanation in B3.' },
      ],
      rubric: [['B3 explains that " CAFETERIA" and "Cafeteria" would become separate groups', 'Teacher check']],
    },
    quiz: null,
  });

  // ---------------------------------------------------------------- Week 12
  C.lessons.push({
    week: 12, unit: 3, id: 'week-12',
    title: 'Combining data',
    summary: 'Appending same-shape tables (including a whole folder of files), merging related tables with the right join kind, and unpivoting wide tables into tidy ones.',
    objectives: [
      'Append tables with the same columns, from separate queries or a folder.',
      'Merge two tables on a key and choose the join kind (left outer, inner, full outer, anti).',
      'Predict and compare row counts for different joins to find unmatched records.',
      'Unpivot a wide table into a tidy, long one.',
    ],
    sections: [
      { heading: 'Append: stacking rows', html: `
        <p><strong>Append</strong> puts the rows of one table under another. Use it when tables have the <em>same columns</em> but different rows: January, February and March sales. Columns are matched by name; a column missing from one table is filled with null.</p>
        <p>Even better, <strong>From Folder</strong> combines every file in a folder in one step and adds a <code>Source.Name</code> column telling you which file each row came from. Drop April's file in the folder, refresh, and it is included.</p>` },
      { heading: 'Merge: adding columns from a related table', html: `
        <p><strong>Merge</strong> joins two tables side by side on a matching key, the same idea as XLOOKUP in week 7 and the same primary key / foreign key relationship. Sales has <code>Product ID</code> (a foreign key); Products has <code>ProductID</code> (its primary key). Merging brings each product's Category and Supplier into the sales rows.</p>` },
      { heading: 'Join kinds', html: `
        <table class="notes-table">
          <thead><tr><th>Join kind</th><th>Keeps</th><th>Use it to</th></tr></thead>
          <tbody>
            <tr><td><strong>Left outer</strong> (default)</td><td>All rows from the first table, matching rows from the second (blank if none)</td><td>Add details without losing any sales</td></tr>
            <tr><td><strong>Inner</strong></td><td>Only rows that match in both</td><td>Keep only sales with a valid product</td></tr>
            <tr><td><strong>Full outer</strong></td><td>Every row from both</td><td>See everything, matched or not</td></tr>
            <tr><td><strong>Right outer</strong></td><td>All rows from the second table</td><td>List every product, with its sales if any</td></tr>
            <tr><td><strong>Left anti</strong></td><td>Rows in the first with <em>no</em> match</td><td>Find orphan records (bad keys)</td></tr>
            <tr><td><strong>Right anti</strong></td><td>Rows in the second with no match</td><td>Find products that never sold</td></tr>
          </tbody>
        </table>
        <div class="callout tip"><strong>Row counts tell the story.</strong> If a left outer join returns 54 rows and an inner join returns 52, then 2 sales have a product key that is not in the catalog. Compare counts every time you merge. Note that Power Query matches text keys case-sensitively: <code>p101</code> does not match <code>P101</code>.</div>` },
      { heading: 'Unpivot: wide to tidy', html: `
        <p>Remember the tidy rule from week 2. Reports often arrive "wide": one column per month. <strong>Unpivot</strong> turns those columns into rows. Select the column(s) to keep (Club), then <strong>Unpivot Other Columns</strong>: you get <code>Club | Attribute | Value</code>, one row per club per month. Rename Attribute to Month and Value to Attendance. Blank cells are dropped. Now a PivotTable or Power BI chart can use Month as a dimension.</p>` },
      { heading: 'Group By', html: `
        <p>After combining, <strong>Group By</strong> summarizes inside the query: Category → Sum of Revenue. It is the same aggregation you did with SUMIFS (week 6) and PivotTables (week 10), now as a repeatable step.</p>` },
    ],
    keyTerms: [
      ['Append', 'Stack rows from tables with the same columns.'],
      ['Merge', 'Join tables side by side on a matching key.'],
      ['Left outer join', 'All rows from the first table plus matches from the second.'],
      ['Inner join', 'Only rows with a match in both tables.'],
      ['Anti join', 'Rows in one table with no match in the other.'],
      ['Unpivot', 'Turn columns into rows to make a wide table tidy.'],
    ],
    practice: [
      {
        id: 'p12a', title: 'Workbook 12A · Combine a folder, then merge',
        intro: 'New query From Folder with all three sales files, named Q1. Apply the cleaning recipe. Then Merge with products.csv on Product ID = ProductID (Left Outer), bringing in Category. Close & Load, and answer the questions.',
        tools: ['query', 'sheets', 'download', 'reset'],
        workbook: {
          sheets: [{ name: 'Answers', rows: [['Question', 'Answer'], ['Rows in Q1 after cleaning'], ['Rows if you used an Inner join instead'], ['Distinct stores after cleaning']], lock: 'A1:A4', widths: { A: 320, B: 100 }, styles: { 'B2:B4': { fill: 'answer' } } }],
          sources: pick(...monthFiles, 'products.csv'),
        },
        questions: [
          { id: 'q1', type: 'query', solution: { name: 'Q1', source: { kind: 'folder', files: monthFiles }, steps: [...cleanSteps(), { type: 'merge', with: { kind: 'file', file: 'products.csv' }, leftKey: 'Product ID', rightKey: 'ProductID', kind: 'leftOuter', columns: ['Category'] }] }, steps: ['removeBlankRows', 'removeDuplicates', 'transform', 'changeType', 'addColumn', 'merge'], loaded: true, points: 6, prompt: 'Folder query Q1, cleaned, merged with products, loaded.' },
          { id: 'q2', type: 'cell', sheet: 'Answers', cell: 'B2', solution: '=COUNTA(Q1!B:B)-1', prompt: 'Rows in Q1.' },
          { id: 'q3', type: 'cell', sheet: 'Answers', cell: 'B3', solution: '=COUNTA(Q1!I:I)-1', prompt: 'Rows with an inner join (every sale whose product is in the catalog).' },
          { id: 'q4', type: 'cell', sheet: 'Answers', cell: 'B4', expect: 3, solution: 3, prompt: 'Distinct stores after cleaning.' },
        ],
      },
      {
        id: 'p12b', title: 'Workbook 12B · Unpivot attendance',
        intro: 'New query on attendance_wide.csv named Attendance. Select Club and Unpivot Other Columns. Rename Attribute to Month and Value to Attendance, change Attendance to Whole Number, then Close & Load.',
        tools: ['query', 'sheets', 'download', 'reset'],
        workbook: { sheets: [startSheet(['attendance_wide.csv has one column per month. Make it tidy with Unpivot.'])], sources: pick('attendance_wide.csv') },
        questions: [
          { id: 'q1', type: 'query', solution: { name: 'Attendance', source: { kind: 'file', file: 'attendance_wide.csv' }, steps: [{ type: 'unpivot', keep: ['Club'] }, { type: 'renameColumn', from: 'Attribute', to: 'Month' }, { type: 'renameColumn', from: 'Value', to: 'Attendance' }, { type: 'changeType', columns: { Attendance: 'whole' } }] }, steps: ['unpivot', 'renameColumn', 'changeType'], loaded: true, points: 4, prompt: 'Attendance query unpivoted, renamed, typed and loaded.' },
          { id: 'q2', type: 'number', prompt: 'How many rows does the tidy table have? (4 clubs × 4 months, minus blanks)', answer: 15 },
          { id: 'q3', type: 'cell', sheet: 'Attendance', cell: 'B1', expect: 'Month', solution: null, prompt: 'Column B of the loaded table is named Month.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-12', title: 'Quarterly category report', points: 25,
      html: `<p>Build one refreshable query that answers "Which product category earned the most in Q1?"</p>
        <ol><li>From Folder: all three monthly files. Name the query <code>Q1Category</code>.</li>
        <li>Clean (blank rows, duplicates, types, Revenue column).</li>
        <li>Merge with <code>products.csv</code> to bring in Category (Left Outer).</li>
        <li>Group By Category, Sum of Revenue as <code>Revenue</code>; sort Revenue largest first. Close &amp; Load.</li>
        <li>On <code>Report</code>, fill B2 (top category), B3 (its revenue), and B4: a left anti merge of the sales against products would return how many rows? Explain in B5 what those rows would mean.</li></ol>`,
      tools: ['query', 'format', 'sheets', 'download', 'reset'],
      workbook: {
        sheets: [{ name: 'Report', rows: [['Question', 'Answer'], ['Top category in Q1'], ['Its revenue'], ['Rows from a left anti merge (sales with no matching product)'], ['What those rows would mean']], formats: { B3: 'currency' }, lock: 'A1:A5', widths: { A: 420, B: 380 }, styles: { 'B2:B5': { fill: 'answer' } } }],
        sources: pick(...monthFiles, 'products.csv'),
      },
      questions: [
        { id: 'a1', type: 'query', solution: { name: 'Q1Category', source: { kind: 'folder', files: monthFiles }, steps: [...cleanSteps(), { type: 'merge', with: { kind: 'file', file: 'products.csv' }, leftKey: 'Product ID', rightKey: 'ProductID', kind: 'leftOuter', columns: ['Category'] }, { type: 'groupBy', by: ['Category'], aggregations: [{ name: 'Revenue', op: 'sum', column: 'Revenue' }] }, { type: 'sortRows', column: 'Revenue', order: 'desc' }] }, steps: ['removeBlankRows', 'removeDuplicates', 'changeType', 'addColumn', 'merge', 'groupBy', 'sortRows'], loaded: true, points: 14, prompt: 'Q1Category query: folder, cleaned, merged, grouped, sorted, loaded.' },
        { id: 'a2', type: 'cell', sheet: 'Report', cell: 'B2', solution: '=Q1Category!A2', points: 3, prompt: 'Top category.' },
        { id: 'a3', type: 'cell', sheet: 'Report', cell: 'B3', solution: '=Q1Category!B2', tolerance: 0.005, points: 3, prompt: 'Its revenue.' },
        { id: 'a4', type: 'cell', sheet: 'Report', cell: 'B4', expect: 0, solution: 0, points: 2, prompt: 'Rows from a left anti merge.' },
        { id: 'a5', type: 'filled', sheet: 'Report', range: 'B5', min: 1, points: 3, prompt: 'Explanation in B5.' },
      ],
    },
    quiz: 'quiz-12',
  });

  // ---------------------------------------------------------------- Week 13
  C.lessons.push({
    week: 13, unit: 3, id: 'week-13',
    title: 'Power BI reports',
    summary: 'From Excel to Power BI: loading clean data, the data model and star schema, relationships, and building a report page with cards, a chart, a table and a slicer.',
    objectives: [
      'Describe what Power BI adds beyond Excel, and name its three views (Report, Table/Data, Model).',
      'Load data into Power BI through Power Query.',
      'Separate fact and dimension tables and connect them with one-to-many relationships (a star schema).',
      'Build a report page with a card, a chart, a table and a slicer.',
      'Rebuild a week 10 PivotTable as a Power BI visual.',
    ],
    sections: [
      { heading: 'Why Power BI?', html: `
        <p>Excel is the best tool for exploring and calculating. <strong>Power BI</strong> is built for <strong>sharing interactive reports</strong>: it handles millions of rows, connects many tables in one model, refreshes on a schedule, and publishes dashboards that people view in a browser or phone without editing them. The skills carry over directly: Power BI has the same Power Query, and its visuals work like PivotTables.</p>
        <div class="callout"><strong>Getting Power BI:</strong> Power BI Desktop is free on Windows. On a Chromebook, sign in to the Power BI service (app.powerbi.com) with a school account, or use a lab PC. If neither is available, the Excel version of this week's lab on this site practices the same skills.</div>` },
      { heading: 'The three views', html: `
        <ul>
          <li><strong>Report view:</strong> the canvas where you build visuals.</li>
          <li><strong>Table view</strong> (Data view): the loaded tables, like a read-only spreadsheet.</li>
          <li><strong>Model view:</strong> boxes for each table with lines showing relationships.</li>
        </ul>
        <p>To load data: <strong>Home › Get data › Text/CSV</strong> (or Folder), then <strong>Transform data</strong> to open Power Query. Everything from weeks 11–12 works here unchanged. Click <strong>Close &amp; Apply</strong> to load.</p>` },
      { heading: 'The data model and the star schema', html: `
        <p>In week 7 you used XLOOKUP to copy product names into the sales sheet. Power BI does not need to copy anything: it keeps the tables separate and <strong>relates</strong> them.</p>
        <ul>
          <li>A <strong>fact table</strong> holds events or transactions with measures: Sales (one row per order line, with Qty and Revenue).</li>
          <li><strong>Dimension tables</strong> describe the things involved: Products, Stores, Dates, Students.</li>
          <li>Each dimension connects to the fact table with a <strong>one-to-many relationship</strong>: the primary key in the dimension (Products[ProductID], the "1" side) to the foreign key in the fact table (Sales[Product ID], the "many" side).</li>
        </ul>
        <p>Drawn out, the fact table sits in the middle with dimensions around it: a <strong>star schema</strong>. In Model view, drag Products[ProductID] onto Sales[Product ID] to create the relationship. Filters flow from the "1" side to the "many" side: choosing Category = Apparel in a slicer filters the Sales rows.</p>` },
      { heading: 'Building the report page', html: `
        <table class="notes-table">
          <thead><tr><th>Visual</th><th>Use it for</th><th>How</th></tr></thead>
          <tbody>
            <tr><td><strong>Card</strong></td><td>One headline number: total revenue, number of orders</td><td>Card visual; drag Revenue to Fields</td></tr>
            <tr><td><strong>Clustered bar/column chart</strong></td><td>Comparing categories</td><td>Category on the Y-axis (or X), Sum of Revenue on Values</td></tr>
            <tr><td><strong>Line chart</strong></td><td>Trend over time</td><td>Order Date (Month) on X-axis, Revenue on Y-axis</td></tr>
            <tr><td><strong>Table / Matrix</strong></td><td>Detail rows, or a pivot-style grid</td><td>Matrix: Category in Rows, Store in Columns, Revenue in Values</td></tr>
            <tr><td><strong>Slicer</strong></td><td>Letting the reader filter</td><td>Slicer visual with Store or Month</td></tr>
          </tbody>
        </table>
        <p>Click a bar in one visual and the others filter to match: this is called <strong>cross-filtering</strong>, and it is the biggest difference from a static Excel chart.</p>
        <p>The Matrix visual with Category in Rows and Sum of Revenue in Values is <em>exactly</em> the PivotTable from week 10. Same aggregation, new tool.</p>` },
      { heading: 'Lab checklist', html: `
        <ol>
          <li>Get data › Folder (the three sales files) › Transform: apply your week 11 recipe. Name it Sales.</li>
          <li>Get data › Text/CSV › products.csv. Name it Products.</li>
          <li>Close &amp; Apply. In Model view, relate Products[ProductID] 1 → ∞ Sales[Product ID].</li>
          <li>Report view: two cards (Total Revenue, number of orders), a bar chart of Revenue by Category, a matrix of Category × Store, and a Store slicer.</li>
          <li>Give the page a title that states the main finding. Save as a .pbix and export a PDF or screenshot.</li>
        </ol>` },
    ],
    keyTerms: [
      ['Power BI', 'Microsoft\'s tool for building and sharing interactive reports.'],
      ['Data model', 'The set of tables and relationships a report is built on.'],
      ['Fact table', 'A table of events or transactions with measures (Sales).'],
      ['Dimension table', 'A table that describes things (Products, Stores, Dates).'],
      ['Star schema', 'A fact table in the middle related one-to-many to dimension tables.'],
      ['Cross-filtering', 'Selecting a value in one visual filters the others.'],
      ['Card', 'A visual that shows one headline number.'],
    ],
    practice: [
      {
        id: 'p13a', title: 'Workbook 13A · Model it (concepts)',
        intro: 'Answer on the page. These are the decisions you make in Power BI\'s Model view.',
        tools: 'basic',
        workbook: null,
        questions: [
          { id: 'q1', type: 'choice', prompt: 'In the store model, which is the fact table?', options: ['Products', 'Sales', 'Stores', 'Dates'], answer: 1 },
          { id: 'q2', type: 'choice', prompt: 'The relationship between Products and Sales is…', options: ['One product to many sales', 'Many products to one sale', 'One-to-one', 'There is no relationship'], answer: 0 },
          { id: 'q3', type: 'match', prompt: 'Match each visual to its best use.', pairs: [['Card', 'One headline number'], ['Line chart', 'Trend over time'], ['Matrix', 'A pivot-style grid of two dimensions'], ['Slicer', 'Letting the reader filter the page']] },
          { id: 'q4', type: 'choice', prompt: 'You click the "Apparel" bar and the other visuals update to show only apparel. This is called…', options: ['Drilling through', 'Cross-filtering', 'Refreshing', 'Unpivoting'], answer: 1 },
          { id: 'q5', type: 'text', prompt: 'Which column in Sales is the foreign key to Products?', accept: ['Product ID', 'ProductID', 'Sales[Product ID]'] },
        ],
      },
      {
        id: 'p13b', title: 'Workbook 13B · The report, built in Excel',
        intro: 'No Power BI today? Build the same report here. The Q1 sales are already loaded. Fill the two cards with formulas, build a pivot of Revenue by Category, and chart it.',
        tools: 'analyze',
        workbook: {
          sheets: [{ name: 'Report', rows: [['Store report, Q1 2026'], [null], ['Total revenue', null, null, 'Orders'], [null]], lock: ['A1', 'A3', 'D3'], widths: { A: 130, B: 100, D: 80, E: 80 }, formats: { B3: 'currency' }, styles: { A1: { bold: true }, B3: { fill: 'answer', bold: true }, E3: { fill: 'answer', bold: true } } }],
          sources: pick(...monthFiles, 'products.csv'),
          queries: [{ id: 'q-Q1', name: 'Q1', source: { kind: 'folder', files: monthFiles }, steps: [...cleanSteps(), { type: 'merge', with: { kind: 'file', file: 'products.csv' }, leftKey: 'Product ID', rightKey: 'ProductID', kind: 'leftOuter', columns: ['Category'] }] }],
        },
        model: (wb) => { wb.loadQuery('q-Q1'); },
        questions: [
          { id: 'q1', type: 'cell', sheet: 'Report', cell: 'B3', solution: '=SUM(Q1!H:H)', requireFormula: true, tolerance: 0.005, points: 2, prompt: 'Card: total revenue (a formula on the Q1 sheet).' },
          { id: 'q2', type: 'cell', sheet: 'Report', cell: 'E3', solution: '=COUNT(Q1!B:B)', requireFormula: true, points: 2, prompt: 'Card: number of orders.' },
          { id: 'q3', type: 'pivot', sheet: 'Q1', solution: { sourceSheet: 'Q1', rows: 'Category', values: 'Revenue', agg: 'sum' }, points: 3, prompt: 'Pivot: Sum of Revenue by Category from the Q1 sheet.' },
          { id: 'q4', type: 'chart', chartType: ['bar', 'column'], solution: { sheet: 'Pivot1', type: 'bar', range: 'A1:B5', title: 'Apparel leads Q1 revenue' }, anyRange: true, requireTitle: true, axisFromZero: true, points: 3, prompt: 'A titled bar or column chart of the pivot.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-13', title: 'Power BI store report', points: 30,
      html: `<p>Build the lab report from the notes in Power BI and turn in evidence. Then answer the model questions here.</p>
        <ol><li><strong>Power BI:</strong> follow the lab checklist (Sales from the folder, Products, a 1-to-many relationship, two cards, a bar chart, a matrix, a slicer, a title that states a finding). Upload a PDF export or screenshot (or the .pbix).</li>
        <li><strong>No Power BI access?</strong> Build the same page in the web workbook (Workbook 13B steps) and upload the downloaded .xlsx instead.</li>
        <li>Answer the questions below; they are graded automatically.</li></ol>`,
      upload: { accept: '.pbix,.pdf,.png,.jpg,.jpeg,.xlsx', label: 'Report file (PDF, screenshot, .pbix or .xlsx)' },
      questions: [
        { id: 'a1', type: 'choice', points: 3, prompt: 'Which table is on the "1" side of the Products–Sales relationship?', options: ['Sales', 'Products', 'Both', 'Neither'], answer: 1 },
        { id: 'a2', type: 'choice', points: 3, prompt: 'A Store slicer is set to "Gym Kiosk". What do the cards show?', options: ['Totals for all stores', 'Totals for Gym Kiosk only', 'Nothing until you refresh', 'An error'], answer: 1 },
        { id: 'a3', type: 'choice', points: 3, prompt: 'Your bar chart shows Category from Products and Revenue from Sales, but every bar is the same total. What is the most likely cause?', options: ['The chart type is wrong', 'There is no relationship between Products and Sales', 'Revenue is text', 'The slicer is on'], answer: 1 },
        { id: 'a4', type: 'match', points: 4, prompt: 'Match each Excel skill to its Power BI equivalent.', pairs: [['PivotTable', 'Matrix visual'], ['XLOOKUP between sheets', 'A relationship in Model view'], ['A single SUM cell on a dashboard', 'Card visual'], ['Filter buttons on a pivot', 'Slicer']] },
      ],
      rubric: [
        ['Data loaded through Power Query with the cleaning recipe', 3],
        ['Products related to Sales one-to-many', 3],
        ['Two cards, a bar chart, a matrix and a slicer', 6],
        ['Title states a finding; visuals readable; axes honest', 5],
      ],
      manualPoints: 17,
    },
    quiz: 'quiz-13',
  });

  // ---------------------------------------------------------------- Week 14
  const piiCols = [
    ['Column', 'Example', 'Classification'],
    ['StudentID', 'S1004', null],
    ['Full Name', 'Ethan Nguyen', null],
    ['Email', 'enguyen@school.org', null],
    ['Home Address', '41 Elm St', null],
    ['Birth Date', '3/9/2008', null],
    ['Zip Code', '02139', null],
    ['Grade Level', '12', null],
    ['Lunch Status', 'Free/Reduced', null],
    ['Favorite Lunch', 'Pizza', null],
    ['Club', 'Robotics', null],
  ];
  const piiAnswers = ['Direct', 'Direct', 'Direct', 'Direct', 'Indirect', 'Indirect', 'Not PII', 'Sensitive', 'Not PII', 'Not PII'];

  C.lessons.push({
    week: 14, unit: 3, id: 'week-14',
    title: 'Data storytelling and ethics',
    summary: 'DAX measures, designing a dashboard around one question, telling a story with data, and the responsibilities that come with personal data: PII, privacy laws like FERPA, aggregation and bias.',
    objectives: [
      'Write simple DAX measures with SUM, COUNTROWS, DISTINCTCOUNT, AVERAGE, DIVIDE and CALCULATE.',
      'Explain the difference between a measure and a calculated column.',
      'Design a dashboard around one question with a clear visual hierarchy.',
      'Identify direct, indirect and sensitive personal data, and decide what should not be shown.',
      'Describe how data and dashboards can cause harm, and how to reduce it.',
    ],
    sections: [
      { heading: 'DAX measures', html: `
        <p><strong>DAX</strong> (Data Analysis Expressions) is Power BI's formula language. Its formulas look like Excel's but work on <em>whole columns and tables</em>, and they recalculate for whatever the visual and slicers are showing.</p>
        <table class="notes-table">
          <thead><tr><th>DAX measure</th><th>Excel equivalent</th></tr></thead>
          <tbody>
            <tr><td><code>Total Revenue = SUM(Sales[Revenue])</code></td><td><code>=SUM(Sales!H:H)</code></td></tr>
            <tr><td><code>Orders = COUNTROWS(Sales)</code></td><td><code>=COUNTA(Sales!B2:B55)</code></td></tr>
            <tr><td><code>Products Sold = DISTINCTCOUNT(Sales[Product ID])</code></td><td>Count of unique values</td></tr>
            <tr><td><code>Avg Order = DIVIDE([Total Revenue], [Orders])</code></td><td><code>=IFERROR(B2/B3, BLANK)</code>: DIVIDE never shows a divide-by-zero error</td></tr>
            <tr><td><code>Cafeteria Revenue = CALCULATE([Total Revenue], Sales[Store] = "Cafeteria")</code></td><td><code>=SUMIFS(…,"Cafeteria")</code></td></tr>
          </tbody>
        </table>
        <p>Create one with <strong>Home › New measure</strong>. Columns are written <code>Table[Column]</code>; measures are written <code>[Measure]</code>. Put a measure in a card and it shows the total; put it in a bar chart by Category and it automatically calculates for each category. That is the power of a measure.</p>
        <p>A <strong>calculated column</strong> (like <code>Revenue = Sales[Qty] * Sales[Unit Price]</code>) is computed once per row and stored, like a helper column. A <strong>measure</strong> is computed on the fly for the current filter. Rule of thumb: if you want to slice <em>by</em> it, make a column; if you want to aggregate it, make a measure.</p>` },
      { heading: 'Designing a dashboard', html: `
        <ul>
          <li><strong>One question per page.</strong> "Where should the store focus next quarter?" Everything on the page should help answer it.</li>
          <li><strong>Visual hierarchy:</strong> the most important number (a card) top-left, where eyes go first; supporting charts next; details (tables) at the bottom.</li>
          <li><strong>Five or six visuals at most</strong>, aligned on a grid, with consistent colors (one color per category on every chart).</li>
          <li><strong>Titles that state findings</strong>, and a short text box with the recommendation.</li>
          <li><strong>Honest scales</strong> (week 9) and clearly labeled filters, so readers know what they are seeing.</li>
        </ul>` },
      { heading: 'Telling the story', html: `
        <p>A data story has three parts: <strong>context</strong> (what question, why it matters), <strong>insight</strong> (what the data shows, with one or two key numbers), and <strong>action</strong> (what should happen next). Lead with the conclusion: "Apparel is 40% of revenue from 10% of orders, so the store should feature apparel at events," then show the evidence.</p>` },
      { heading: 'Personal data and PII', html: `
        <p><strong>Personally identifiable information (PII)</strong> is data that can identify a person.</p>
        <ul>
          <li><strong>Direct identifiers</strong> identify someone on their own: full name, student ID, email, home address, phone, Social Security number.</li>
          <li><strong>Indirect (quasi) identifiers</strong> identify someone when combined: birth date + zip code + gender identifies most Americans uniquely.</li>
          <li><strong>Sensitive data</strong> could cause harm if revealed: health information, free/reduced lunch status, disability, discipline records, immigration status.</li>
        </ul>
        <p>In the US, <strong>FERPA</strong> protects student education records: schools generally cannot share identifiable records without consent. Other laws protect health data (HIPAA), children's data online (COPPA), and, in Europe, all personal data (GDPR). You do not need to memorize the laws; you need the habit of asking "should this be here?"</p>` },
      { heading: 'Reducing harm', html: `
        <ul>
          <li><strong>Data minimization:</strong> only collect and show what the question needs. A dashboard about club attendance does not need student names. Remove the column in Power Query so it never reaches the report.</li>
          <li><strong>Aggregate:</strong> show counts and averages, not individual rows.</li>
          <li><strong>Small-group suppression:</strong> if a group has fewer than about 5 people (e.g. "Grade 11 students in Web Design with free lunch: 1"), hide or combine it; a small count can identify someone.</li>
          <li><strong>Access control:</strong> who can see the report? Publishing to the web makes it public.</li>
          <li><strong>Bias and fairness:</strong> data reflects how it was collected (week 3). Ask who is missing and who could be harmed by a decision made from this dashboard.</li>
        </ul>
        <div class="callout"><strong>Discussion:</strong> a dashboard ranks students by attendance and is shown on the hallway TV. Every number is accurate. What is wrong with it, and how would you redesign it?</div>` },
    ],
    keyTerms: [
      ['DAX', 'Power BI\'s formula language.'],
      ['Measure', 'A DAX calculation evaluated for the current filters (e.g. Total Revenue).'],
      ['Calculated column', 'A column computed once per row and stored in the table.'],
      ['PII', 'Personally identifiable information.'],
      ['Quasi-identifier', 'Data that identifies a person when combined with other data.'],
      ['FERPA', 'US law protecting the privacy of student education records.'],
      ['Data minimization', 'Collecting and showing only the data a purpose requires.'],
    ],
    practice: [
      {
        id: 'p14a', title: 'Workbook 14A · DAX by translation',
        intro: 'Each question shows an Excel formula you already know. Pick the DAX measure that does the same job.',
        tools: 'basic',
        workbook: null,
        questions: [
          { id: 'q1', type: 'choice', prompt: 'Excel: =SUM(Sales!H:H). DAX:', options: ['Total = SUMX(Sales)', 'Total = SUM(Sales[Revenue])', 'Total = Sales[Revenue] + 1', 'Total = COUNT(Sales)'], answer: 1 },
          { id: 'q2', type: 'choice', prompt: 'Excel: =SUMIFS(Revenue, Store, "Cafeteria"). DAX:', options: ['CALCULATE([Total Revenue], Sales[Store] = "Cafeteria")', 'FILTER("Cafeteria")', 'SUMIF(Sales[Store], "Cafeteria")', 'COUNTROWS(Sales[Store])'], answer: 0 },
          { id: 'q3', type: 'choice', prompt: 'Which DAX function divides safely, returning BLANK instead of an error when dividing by zero?', options: ['DIV', 'QUOTIENT', 'DIVIDE', 'SAFE'], answer: 2 },
          { id: 'q4', type: 'choice', prompt: 'You want to slice a chart BY "Order Size" (Small/Large). Measure or calculated column?', options: ['Measure', 'Calculated column'], answer: 1 },
          { id: 'q5', type: 'text', prompt: 'Write the DAX function that counts the rows of a table.', accept: ['COUNTROWS', 'COUNTROWS()', 'COUNTROWS(Sales)'] },
        ],
      },
      {
        id: 'p14b', title: 'Workbook 14B · PII audit and minimization',
        intro: 'Classify each column: Direct, Indirect, Sensitive or Not PII. Then open Power Query: the Clubs query feeds a club-attendance dashboard. Remove every column the dashboard does not need (keep only Grade and Club) and Close & Load.',
        tools: ['query', 'sheets', 'download', 'reset'],
        workbook: {
          sheets: [{ name: 'Audit', rows: piiCols, lock: ['A1:C1', 'A2:B11'], validations: [{ range: 'C2:C11', type: 'list', values: ['Direct', 'Indirect', 'Sensitive', 'Not PII'] }], widths: { A: 130, B: 150, C: 130 }, styles: { 'C2:C11': { fill: 'answer' } } }],
          sources: { 'club_roster.csv': 'StudentID,Full Name,Email,Grade,Lunch Status,Club\nS1001,Maya Patel,mpatel@school.org,12,Paid,Robotics\nS1002,Luis Garcia,lgarcia@school.org,11,Free/Reduced,Debate\nS1003,Aisha Bello,abello@school.org,12,Paid,Robotics\nS1004,Ethan Nguyen,enguyen@school.org,12,Free/Reduced,Art\nS1005,Grace Kim,gkim@school.org,11,Paid,Debate\nS1006,Noah Smith,nsmith@school.org,12,Paid,Coding\nS1007,Sofia Rossi,srossi@school.org,12,Free/Reduced,Art' },
          queries: [{ id: 'q-Clubs', name: 'Clubs', source: { kind: 'file', file: 'club_roster.csv' }, steps: [] }],
        },
        questions: [
          { id: 'q1', type: 'range', sheet: 'Audit', range: 'C2:C11', expect: piiAnswers.map((x) => [x]), points: 5, prompt: 'Classify all ten columns.' },
          { id: 'q2', type: 'query', solution: { name: 'Clubs', source: { kind: 'file', file: 'club_roster.csv' }, steps: [{ type: 'keepColumns', columns: ['Grade', 'Club'] }] }, steps: [['keepColumns', 'removeColumns']], loaded: true, points: 2, prompt: 'Clubs query removes the unneeded columns and is loaded.' },
          { id: 'q3', type: 'cell', sheet: 'Clubs', cell: 'C1', expect: null, solution: null, points: 3, prompt: 'The loaded Clubs sheet has only two columns (nothing in C1).' },
          { id: 'q4', type: 'cell', sheet: 'Clubs', cell: 'A1', expect: 'Grade', solution: null, points: 1, prompt: 'Column A of Clubs is Grade.' },
        ],
      },
    ],
    assignment: {
      id: 'hw-14', title: 'Final dashboard project', points: 50, category: 'project',
      html: `<p>Answer one question for the school store with a one-page dashboard, and present it in week 15.</p>
        <p><strong>Your question</strong> must be measurable (week 6), for example: "Which store location and category should get more shelf space next quarter?"</p>
        <ol><li>Data: a query that combines all three monthly files, cleans them, and merges Category from products.csv. No names or other PII on the page.</li>
        <li>At least one card (headline number), two charts of different types, and a PivotTable or matrix.</li>
        <li>A title that states your answer and a 3-sentence recommendation (context, insight, action).</li>
        <li>Build it in <strong>Power BI</strong> (upload a PDF/screenshot or .pbix) <em>or</em> in the web workbook below (submit it here or upload the .xlsx). The workbook parts are checked automatically; the design and story are scored with the rubric.</li></ol>`,
      tools: 'all',
      upload: { accept: '.pbix,.pdf,.png,.jpg,.jpeg,.xlsx,.pptx', label: 'Optional: Power BI file, PDF, screenshot or slides' },
      workbook: {
        sheets: [{ name: 'Dashboard', rows: [['Dashboard title (state your answer)'], [null], ['My question'], [null], ['Headline number'], [null], ['Recommendation (context, insight, action)']], lock: ['A1', 'A3', 'A5', 'A7'], widths: { A: 380, B: 160 }, styles: { A2: { fill: 'answer' }, A4: { fill: 'answer' }, B5: { fill: 'answer' }, A8: { fill: 'answer' } } }],
        sources: pick(...monthFiles, 'products.csv'),
      },
      model: (wb) => {
        wb.setInput('Dashboard', 'A2', 'Cafeteria apparel should get more space', { force: true });
        wb.setInput('Dashboard', 'A4', 'Which store and category earned the most revenue in Q1 2026?', { force: true });
        wb.setInput('Dashboard', 'A8', 'Context, insight and action go here.', { force: true });
      },
      questions: [
        { id: 'a1', type: 'query', solution: { name: 'Q1', source: { kind: 'folder', files: monthFiles }, steps: [...cleanSteps(), { type: 'merge', with: { kind: 'file', file: 'products.csv' }, leftKey: 'Product ID', rightKey: 'ProductID', kind: 'leftOuter', columns: ['Category'] }] }, steps: ['removeBlankRows', 'removeDuplicates', 'changeType', 'addColumn', 'merge'], loaded: true, points: 8, prompt: 'Combined, cleaned, merged query, loaded.' },
        { id: 'a2', type: 'pivot', sheet: 'Q1', solution: { sourceSheet: 'Q1', rows: 'Category', columns: 'Store', values: 'Revenue', agg: 'sum' }, points: 6, prompt: 'A PivotTable of Revenue by Category (and Store).' },
        { id: 'a3', type: 'cell', sheet: 'Dashboard', cell: 'B5', solution: '=SUM(Q1!H:H)', requireFormula: true, tolerance: 0.005, points: 4, prompt: 'Headline number with a formula (e.g. total Q1 revenue).' },
        { id: 'a4', type: 'chart', chartType: ['bar', 'column'], solution: { sheet: 'Pivot1', type: 'column', range: 'A1:B5', title: 'Apparel leads revenue' }, anyRange: true, requireTitle: true, axisFromZero: true, points: 4, prompt: 'A titled bar or column chart.' },
        { id: 'a5', type: 'chart', chartType: ['line', 'pie', 'scatter'], solution: { sheet: 'Pivot1', type: 'pie', range: 'A1:B5', title: 'Share of revenue by category' }, anyRange: true, requireTitle: true, points: 4, prompt: 'A second chart of a different type (line, pie or scatter), titled.' },
        { id: 'a6', type: 'filled', sheet: 'Dashboard', range: 'A2', points: 1, prompt: 'Dashboard title written in A2.' },
        { id: 'a7', type: 'filled', sheet: 'Dashboard', range: 'A4', points: 1, prompt: 'Measurable question written in A4.' },
        { id: 'a8', type: 'filled', sheet: 'Dashboard', range: 'A8', points: 2, prompt: 'Recommendation written in A8.' },
      ],
      rubric: [
        ['The question is measurable and the title answers it', 4],
        ['Visuals fit their questions; honest axes; consistent colors; clear hierarchy', 6],
        ['Recommendation has context, insight and action, backed by numbers', 5],
        ['No PII on the page; small groups not exposed', 3],
        ['Presentation: explains and defends the decisions (week 15)', 2],
      ],
      manualPoints: 20,
    },
    quiz: 'quiz-14',
  });

  // ---------------------------------------------------------------- Week 15
  C.lessons.push({
    week: 15, unit: 3, id: 'week-15',
    title: 'Synthesis and final exam',
    summary: 'Presenting and defending a dashboard, connecting every unit back to the data-to-decision cycle, and a review of Unit 3 before Friday\'s exam.',
    objectives: [
      'Present a dashboard and defend the decisions behind it: the question, the data, the cleaning, the visuals and what was left out.',
      'Trace one decision through the whole data-to-decision cycle using skills from every unit.',
      'Review Power Query, combining data, Power BI modeling, DAX and data ethics for Exam 3.',
    ],
    sections: [
      { heading: 'Presenting your dashboard', html: `
        <p>Each presentation is 3 minutes plus 2 minutes of questions. Use this structure:</p>
        <ol>
          <li><strong>The question</strong> and why it matters to the store.</li>
          <li><strong>The data:</strong> where it came from, what you cleaned (show Applied Steps), and how the tables relate.</li>
          <li><strong>The answer:</strong> walk the page top-left to bottom-right, one finding per visual.</li>
          <li><strong>The recommendation:</strong> what the store should do, and how it would measure whether it worked.</li>
          <li><strong>The decisions:</strong> one chart you chose not to use and why; what data you left out on purpose (PII).</li>
        </ol>
        <p>Expect questions like "How do you know there were no duplicates?", "Why a bar chart instead of a pie?", "What would change your recommendation?"</p>` },
      { heading: 'The whole course in one cycle', html: `
        <table class="notes-table">
          <thead><tr><th>Step (week 1)</th><th>Skills you now have</th></tr></thead>
          <tbody>
            <tr><td><strong>Ask</strong></td><td>Measurable questions, metrics and KPIs (week 6)</td></tr>
            <tr><td><strong>Collect</strong></td><td>Sources, bias, quality dimensions (week 3); CSV vs Excel and importing (week 4); keys and related tables (week 7)</td></tr>
            <tr><td><strong>Organize</strong></td><td>Types and tidy tables (week 2); Tables and validation (week 5); cleaning (week 8); Power Query (weeks 11–12)</td></tr>
            <tr><td><strong>Analyze</strong></td><td>Formulas (week 3); conditional functions (week 6); lookups (week 7); PivotTables (week 10); DAX (week 14)</td></tr>
            <tr><td><strong>Decide</strong></td><td>Honest charts (week 9); dashboards and storytelling (weeks 13–14)</td></tr>
            <tr><td><strong>Measure and repeat</strong></td><td>Refreshable queries and reports (weeks 11–13); ethics of what you measure (week 14)</td></tr>
          </tbody>
        </table>` },
      { heading: 'Unit 3 review checklist', html: `
        <p>Exam 3 is Friday and covers weeks 11–15 only. You should be able to:</p>
        <ul class="checklist-review">
          <li>Explain ETL and why Power Query's applied steps make cleaning repeatable and auditable.</li>
          <li>Build a query: blank rows, duplicates, trim/capitalize, change type, custom column, Close &amp; Load; change the source and refresh.</li>
          <li>Choose append vs merge; pick a join kind and predict row counts; use a left anti join to find orphans; unpivot a wide table; Group By.</li>
          <li>Describe a star schema: fact vs dimension tables, one-to-many relationships, filter direction, cross-filtering.</li>
          <li>Pick visuals (card, bar, line, matrix, slicer) and design a page around one question.</li>
          <li>Read and write simple DAX: SUM, COUNTROWS, DISTINCTCOUNT, DIVIDE, CALCULATE; measure vs calculated column.</li>
          <li>Classify direct, indirect and sensitive data; apply minimization, aggregation and small-group suppression.</li>
        </ul>` },
    ],
    keyTerms: [
      ['Data-to-decision cycle', 'Ask, collect, organize, analyze, decide, measure.'],
      ['Audit trail', 'A record of what was done to data (Applied Steps).'],
      ['Reproducible', 'Anyone can re-run the same steps on the same data and get the same result.'],
      ['Single source of truth', 'One agreed, maintained place a figure comes from, so reports do not disagree.'],
      ['Data story', 'A headline finding, the evidence that supports it, and the decision it informs.'],
    ],
    practice: [
      {
        id: 'p15a', title: 'Workbook 15A · End-to-end review',
        intro: 'One last full pass: build a query that combines January and February, keeps only Cafeteria rows, and groups Revenue by Product ID. Name it Review and load it.',
        tools: ['query', 'sheets', 'download', 'reset'],
        workbook: { sheets: [startSheet(['Power Query › New query › From Folder (sales_jan.csv and sales_feb.csv). Clean, filter Store = Cafeteria, Group By Product ID (Sum of Revenue).'])], sources: pick('sales_jan.csv', 'sales_feb.csv') },
        questions: [
          { id: 'q1', type: 'query', solution: { name: 'Review', source: { kind: 'folder', files: ['sales_jan.csv', 'sales_feb.csv'] }, steps: [...cleanSteps(), { type: 'filterRows', column: 'Store', op: '=', value: 'Cafeteria' }, { type: 'groupBy', by: ['Product ID'], aggregations: [{ name: 'Revenue', op: 'sum', column: 'Revenue' }] }] }, steps: ['removeBlankRows', 'removeDuplicates', 'transform', 'filterRows', 'groupBy'], loaded: true, points: 5, prompt: 'Review query: folder of two files, cleaned, filtered to Cafeteria, grouped, loaded.' },
          { id: 'q2', type: 'choice', prompt: 'Why must Trim/Capitalize come before the Cafeteria filter?', options: ['Filters only work on numbers', 'Otherwise " CAFETERIA" rows would not match "Cafeteria" and would be dropped', 'Power Query sorts steps automatically', 'It does not matter'], answer: 1 },
        ],
      },
    ],
    assignment: {
      id: 'hw-15', title: 'Course reflection', points: 15,
      html: `<p>Write a short reflection (a few sentences per prompt). It will be read by your teacher, so write in full sentences.</p>`,
      workbook: {
        sheets: [{ name: 'Reflection', rows: [['Prompt', 'Response'], ['Which skill from this course will you use first outside school, and for what?'], ['Describe a time this semester when the data was wrong or messy. How did you find out?'], ['What is one decision you think should NOT be made from data alone? Why?'], ['What would you add to your dashboard with one more week?']], lock: 'A1:A5', widths: { A: 420, B: 520 }, styles: { 'B2:B5': { fill: 'answer' } } }],
      },
      questions: [
        { id: 'a1', type: 'filled', sheet: 'Reflection', range: 'B2:B5', min: 4, points: 8, prompt: 'All four prompts answered.' },
      ],
      rubric: [['Specific examples from the course', 4], ['Thoughtful answer about the limits of data', 3]],
      manualPoints: 7,
    },
    quiz: 'quiz-15',
    exam: 'exam-3',
  });

  if (isNode) module.exports = C;
})(typeof globalThis !== 'undefined' ? globalThis : this);
