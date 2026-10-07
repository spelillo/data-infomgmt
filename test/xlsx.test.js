'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { F, W, X, G, wb, allAssessments, findAssessment, patchXlsx } = require('./helpers');
const { addPivot, addQueries } = require('./excel-parts');

const FIXTURE = path.join(__dirname, 'fixtures', 'excel-structures.xlsx');
const v = (book, sheet, addr) => { const x = book.getValue(sheet, addr); return F.isError(x) ? x.code : x; };

test('zip: stored archives round-trip and CRC32 matches the standard check value', async () => {
  assert.equal(X.crc32(new TextEncoder().encode('123456789')), 0xcbf43926);
  const bytes = X.zipStore([{ name: 'a.txt', data: new TextEncoder().encode('hello') }, { name: 'dir/b.txt', data: new Uint8Array([1, 2, 3]) }]);
  const files = await X.unzip(bytes);
  assert.equal(new TextDecoder().decode(files['a.txt']), 'hello');
  assert.deepEqual([...files['dir/b.txt']], [1, 2, 3]);
});

test('reading a non-Excel file fails with a helpful message', async () => {
  await assert.rejects(X.read(new TextEncoder().encode('not a zip at all')), /not a valid \.xlsx/);
  const zip = X.zipStore([{ name: 'hello.txt', data: new Uint8Array([1]) }]);
  await assert.rejects(X.read(zip), /does not look like an Excel workbook/);
});

test('values, formulas, formats, bold, widths and names survive write → read', async () => {
  const book = W.Workbook.fromSpec({
    sheets: [{
      name: 'Data',
      rows: [['Item', 'Qty', 'Price', 'Total', 'When', 'Code', 'Ok'], ['Pens', 2, 1.5, '=B2*C2', '2026-01-05', "'00123", true], ['Ink', 1, 9, '=B3*C3', '2026-02-01', "'007", false]],
      formats: { 'C2:D3': 'currency', 'E2:E3': 'date' },
      styles: { 'A1:G1': { bold: true } },
      widths: { A: 140 },
    }, { name: 'Calc & Notes', rows: [['=SUM(Data!D2:D3)', '=IFERROR(1/0,"none")', '=1/0']] }],
    names: { Prices: 'Data!C2:C3' },
  });
  const back = await X.read(X.write(book));
  assert.deepEqual(back.sheets.map((s) => s.name), ['Data', 'Calc & Notes']);
  assert.equal(back.getInput('Data', 'D2'), '=B2*C2');
  assert.equal(v(back, 'Data', 'D3'), 9);
  assert.equal(back.getFormat('Data', F.parseAddress('C2')), 'currency');
  assert.equal(back.getFormat('Data', F.parseAddress('E2')), 'date');
  assert.equal(v(back, 'Data', 'E2'), F.dateToSerial(2026, 1, 5));
  assert.equal(v(back, 'Data', 'F2'), '00123');
  assert.equal(v(back, 'Data', 'G2'), true);
  assert.equal(back.getCell('Data', 'A1').style.bold, true);
  assert.equal(back.getSheet('Data').widths.A, 140);
  assert.equal(v(back, 'Calc & Notes', 'A1'), 12);
  assert.equal(v(back, 'Calc & Notes', 'B1'), 'none');
  assert.equal(v(back, 'Calc & Notes', 'C1'), '#DIV/0!');
  assert.ok(back.names.Prices);
});

test('tables, data validation and conditional formatting survive write → read', async () => {
  const book = wb({ S: [['Item', 'Qty'], ['Pens', 2], ['Ink', 1]] });
  book.addTable('S', 'A1:B3', 'Orders');
  book.addValidation('S', 'B2:B20', { type: 'whole', min: 0, max: 50 });
  book.addValidation('S', 'A2:A20', { type: 'list', values: ['Pens', 'Ink', 'Paper'] });
  book.addCondFormat('S', 'B2:B3', { op: '>', value: 1, style: 'green' });
  book.addCondFormat('S', 'A2:A3', { op: 'contains', value: 'en', style: 'yellow' });
  book.addCondFormat('S', 'A2:A3', { op: 'duplicate', style: 'red' });
  book.setInput('S', 'D1', '=SUM(Orders[Qty])');
  const back = await X.read(X.write(book));
  assert.deepEqual(back.tables.map((t) => [t.name, t.ref]), [['Orders', 'A1:B3']]);
  assert.equal(v(back, 'S', 'D1'), 3);
  assert.deepEqual(back.getSheet('S').validations, [
    { range: 'B2:B20', type: 'whole', min: 0, max: 50 },
    { range: 'A2:A20', type: 'list', values: ['Pens', 'Ink', 'Paper'] },
  ]);
  assert.deepEqual(back.getSheet('S').condFormats, [
    { range: 'B2:B3', op: '>', value: 1, style: 'green' },
    { range: 'A2:A3', op: 'contains', value: 'en', style: 'yellow' },
    { range: 'A2:A3', op: 'duplicate', style: 'red' },
  ]);
});

test('a table whose header is not unique text is left out rather than corrupting the file', async () => {
  const book = wb({ S: [['Item', 'Item'], ['a', 'b']] });
  book.tables.push({ name: 'Dup', sheet: 'S', ref: 'A1:B2' });
  const files = await X.unzip(X.write(book));
  assert.ok(!Object.keys(files).some((k) => k.startsWith('xl/tables/')));
});

test('reads a file made by another spreadsheet program (openpyxl fixture)', async () => {
  const book = await X.read(fs.readFileSync(FIXTURE));
  assert.deepEqual(book.sheets.map((s) => s.name), ['Sales', 'Lists', 'Summary']);
  assert.equal(v(book, 'Sales', 'F2'), 96.75);
  assert.deepEqual(book.tables, [{ name: 'SalesTbl', sheet: 'Sales', ref: 'A1:D6' }]);
  const val = book.getSheet('Sales').validations;
  assert.deepEqual(val.find((x) => x.range === 'A2:A50').values, ['East', 'West', 'North', 'South']);
  assert.deepEqual(val.find((x) => x.range === 'B2:B50').values, ['Pens', 'Paper', 'Ink']); // list from a range on another sheet
  assert.deepEqual(val.find((x) => x.range === 'C2:C50'), { range: 'C2:C50', type: 'whole', min: 0, max: 100 });
  const cf = book.getSheet('Sales').condFormats;
  assert.deepEqual(cf.find((x) => x.range === 'D2:D6'), { range: 'D2:D6', op: '>', value: 20, style: 'green' });
  assert.deepEqual(cf.find((x) => x.range === 'B2:B6'), { range: 'B2:B6', op: 'duplicate', style: 'red' });
  assert.deepEqual(cf.find((x) => x.range === 'A2:A6'), { range: 'A2:A6', op: '=', value: 'East', style: 'yellow' });
  const [bar, line] = book.charts;
  assert.equal(bar.type, 'column');
  assert.equal(bar.sheet, 'Summary');
  assert.equal(bar.range, 'A1:B4');
  assert.equal(bar.title, 'East leads revenue');
  assert.equal(bar.axisMin, 0);
  assert.equal(line.type, 'line');
  assert.equal(line.dataSheet, 'Sales');
  assert.deepEqual(book.chartData(line).categories, ['East', 'West', 'East', 'North', 'West']);
});

test('reads charts Excel writes with c: prefixes, bar direction and auto titles', async () => {
  const base = X.write(wb({ S: [['Cat', 'Rev'], ['A', 1], ['B', 2]] }));
  const chartXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Rev</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>S!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>S!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:barChart><c:catAx/><c:valAx><c:scaling><c:orientation val="minMax"/><c:min val="0.5"/></c:scaling></c:valAx></c:plotArea></c:chart></c:chartSpace>';
  const bytes = await patchXlsx(base, {
    'xl/charts/chart1.xml': chartXml,
    'xl/drawings/drawing1.xml': '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
    'xl/drawings/_rels/drawing1.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
    'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
  });
  const book = await X.read(bytes);
  assert.deepEqual(book.charts.map(({ type, range, title, axisMin }) => ({ type, range, title, axisMin })), [{ type: 'bar', range: 'A1:B3', title: 'Rev', axisMin: 0.5 }]);
});

test('reads PivotTables: fields, summary function and date grouping', async () => {
  const src = wb({ Sales: [['Date', 'Region', 'Revenue'], ['2026-01-05', 'East', 10], ['2026-02-05', 'West', 20]], Pivot: [['Row Labels', 'Sum of Revenue']] });
  let bytes = X.write(src);
  const sheets = src.sheets.map((s) => s.name);
  // Excel 365 auto-groups a date field into Years/Quarters/Months fields
  bytes = await addPivot(bytes, sheets, { sheet: 'Pivot', location: 'A3:B6', source: { sheet: 'Sales', ref: 'A1:C3' }, fields: ['Date', 'Region', 'Revenue'], rows: [4, 0], data: { fld: 2 }, groups: [{ name: 'Date', base: 0, groupBy: 'months' }, { name: 'Quarters', base: 0, groupBy: 'quarters' }, { name: 'Years', base: 0, groupBy: 'years' }] });
  bytes = await addPivot(bytes, sheets, { sheet: 'Pivot', location: 'E3:G6', source: { sheet: 'Sales', ref: 'A1:C3' }, fields: ['Date', 'Region', 'Revenue'], rows: [1], cols: [0], data: { fld: 2, subtotal: 'average' } });
  const book = await X.read(bytes);
  const [p1, p2] = book.pivots;
  assert.equal(p1.rows, 'Date');
  assert.equal(p1.rowsGroup, 'month');
  assert.equal(p1.values, 'Revenue');
  assert.equal(p1.agg, 'sum');
  assert.equal(p1.targetSheet, 'Pivot');
  assert.equal(p1.targetCell, 'A3');
  assert.equal(p1.sourceSheet, 'Sales');
  assert.equal(p2.rows, 'Region');
  assert.equal(p2.columns, 'Date');
  assert.equal(p2.agg, 'average');
});

test('a PivotTable built on an Excel Table resolves the table to its range', async () => {
  const src = wb({ Sales: [['Region', 'Revenue'], ['East', 10]], P: [] });
  src.addTable('Sales', 'A1:B2', 'SalesTbl');
  const bytes = await addPivot(X.write(src), ['Sales', 'P'], { sheet: 'P', location: 'A1:B3', source: { name: 'SalesTbl' }, fields: ['Region', 'Revenue'], rows: [0], data: { fld: 1, subtotal: 'count' } });
  const [p] = (await X.read(bytes)).pivots;
  assert.deepEqual([p.sourceSheet, p.sourceRange, p.agg], ['Sales', 'A1:B2', 'count']);
});

const M_SALES = `section Section1;

shared Jan = let
    Source = Csv.Document(File.Contents("C:\\Users\\student\\Downloads\\sales_jan.csv"),[Delimiter=",", Columns=6, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    // remove the empty lines
    #"Removed Blank Rows" = Table.SelectRows(#"Promoted Headers", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    #"Removed Duplicates" = Table.Distinct(#"Removed Blank Rows"),
    #"Trimmed Text" = Table.TransformColumns(#"Removed Duplicates",{{"Store", Text.Trim, type text}}),
    #"Capitalized Each Word" = Table.TransformColumns(#"Trimmed Text",{{"Store", Text.Proper, type text}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Capitalized Each Word",{{"Order ID", Int64.Type}, {"Order Date", type date}, {"Qty", Int64.Type}, {"Unit Price", Currency.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Revenue", each [Qty]*[Unit Price], Currency.Type)
in
    #"Added Custom";

shared #"Store Totals" = let
    Source = Jan,
    #"Filtered Rows" = Table.SelectRows(Source, each ([Store] <> "Gym")),
    #"Grouped Rows" = Table.Group(#"Filtered Rows", {"Store"}, {{"Revenue", each List.Sum([Revenue]), type number}}),
    #"Sorted Rows" = Table.Sort(#"Grouped Rows",{{"Revenue", Order.Descending}})
in
    #"Sorted Rows";

shared Products = let
    Source = Excel.CurrentWorkbook(){[Name="ProductsTbl"]}[Content],
    #"Merged Queries" = Table.NestedJoin(Source, {"ProductID"}, Jan, {"Product ID"}, "Jan", JoinKind.LeftAnti),
    #"Unpivoted Other Columns" = Table.UnpivotOtherColumns(#"Merged Queries", {"ProductID"}, "Attribute", "Value"),
    #"Renamed Columns" = Table.RenameColumns(#"Unpivoted Other Columns",{{"Value", "Detail"}})
in
    #"Renamed Columns";`;

test('M code maps onto Applied Steps', () => {
  const lets = X.mLet('let\n  A = 1,\n  #"B c" = Table.Distinct(A, {"x"})\nin\n  #"B c"');
  assert.deepEqual(lets.map((l) => l.name), ['A', 'B c']);
  const steps = (expr) => X.mSteps(expr).map((s) => s.type);
  assert.deepEqual(steps('Table.RemoveColumns(x,{"A","B"})'), ['removeColumns']);
  assert.deepEqual(steps('Table.SelectColumns(x,{"A"})'), ['keepColumns']);
  assert.deepEqual(steps('Table.RenameColumns(x,{{"A","B"},{"C","D"}})'), ['renameColumn', 'renameColumn']);
  assert.deepEqual(steps('Table.ReplaceValue(x,"N/A","",Replacer.ReplaceText,{"Notes"})'), ['replaceValues']);
  assert.deepEqual(steps('Table.SplitColumn(x, "Name", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv), {"Last", "First"})'), ['splitColumn']);
  assert.deepEqual(steps('Table.RemoveRowsWithErrors(x, {"Qty"})'), ['removeErrors']);
  assert.deepEqual(steps('Table.FirstN(x,5)'), ['keepTopRows']);
  assert.deepEqual(steps('Table.FillDown(x,{"Region"})'), ['fillDown']);
  assert.deepEqual(steps('Table.Combine({Jan, Feb})'), ['append']);
  assert.deepEqual(steps('Table.PromoteHeaders(Source)'), []);
  const f = X.mSteps('Table.SelectRows(x, each Text.Contains([Store], "Caf"))')[0];
  assert.deepEqual([f.column, f.op, f.value], ['Store', 'contains', 'Caf']);
  const g = X.mSteps('Table.SelectRows(x, each [Qty] >= 3 and [Qty] < 9)')[0];
  assert.deepEqual([g.column, g.op, g.value], ['Qty', '>=', 3]);
  const t = X.mSteps('Table.TransformColumnTypes(x,{{"Qty", Int64.Type}, {"When", type date}})')[0];
  assert.deepEqual(t.columns, { Qty: 'whole', When: 'date' });
  assert.deepEqual(X.mSource('Folder.Files("C:\\\\data\\\\Q1")'), { kind: 'folder', files: [], path: 'C:\\\\data\\\\Q1' });
});

test('reads Power Query queries from the DataMashup and links them to their loaded tables', async () => {
  const src = wb({ Jan: [['Order ID', 'Store'], [1, 'Cafeteria'], [2, 'Library']], Products: [['ProductID', 'Name'], ['P1', 'Pens']], Totals: [['Store', 'Revenue'], ['Cafeteria', 9]] });
  src.addTable('Products', 'A1:B2', 'ProductsTbl');
  const bytes = await addQueries(X.write(src), ['Jan', 'Products', 'Totals'], M_SALES, [{ query: 'Jan', sheet: 'Jan', ref: 'A1:B3' }, { query: 'Store Totals', sheet: 'Totals', ref: 'A1:B2' }]);
  const book = await X.read(bytes);
  const byName = Object.fromEntries(book.queries.map((q) => [q.name, q]));
  assert.deepEqual(Object.keys(byName), ['Jan', 'Store Totals', 'Products']);
  assert.deepEqual(byName.Jan.source, { kind: 'file', file: 'sales_jan.csv' });
  assert.deepEqual(byName.Jan.steps.map((s) => s.type), ['removeBlankRows', 'removeDuplicates', 'transform', 'transform', 'changeType', 'addColumn']);
  assert.deepEqual(byName.Jan.steps.slice(2, 4).map((s) => s.op), ['trim', 'proper']);
  assert.equal(byName.Jan.loadedSheet, 'Jan');
  assert.equal(byName.Jan.loadedRows, 2);
  assert.deepEqual(byName['Store Totals'].source, { kind: 'query', name: 'Jan' });
  assert.deepEqual(byName['Store Totals'].steps.map((s) => s.type), ['filterRows', 'groupBy', 'sortRows']);
  assert.equal(byName['Store Totals'].steps[0].op, '<>');
  assert.equal(byName['Store Totals'].steps[2].order, 'desc');
  assert.equal(byName['Store Totals'].loadedSheet, 'Totals');
  assert.deepEqual(byName.Products.source, { kind: 'sheet', sheet: 'Products', range: 'A1:B2', table: 'ProductsTbl' });
  assert.deepEqual(byName.Products.steps.map((s) => s.type), ['merge', 'unpivot', 'renameColumn']);
  assert.equal(byName.Products.steps[0].kind, 'leftAnti');
  // connection-only queries explain themselves; loaded ones return the loaded table
  assert.match(book.runQuery(byName.Products).error, /connection only/);
  assert.equal(book.runQuery(byName.Jan).table.rows.length, 2);
});

// ----- end to end: files made in Excel are graded like work done on the site --

const WRITABLE = new Set(['cell', 'range', 'cellType', 'filled', 'format', 'table', 'validation', 'condFormat', 'sorted', 'choice', 'number', 'text', 'match']);

test('every answer key survives a trip through .xlsx and still earns full marks', async () => {
  for (const [label, a] of allAssessments()) {
    const qs = (a.questions || []).filter((q) => WRITABLE.has(q.type));
    if (!qs.some(G.isWorkbookQuestion)) continue;
    const back = await X.read(X.write(G.buildKey(a)));
    back.sources = { ...((a.workbook && a.workbook.sources) || {}) };
    // charts, PivotTables and queries are not written by the site (students
    // build those in Excel), so only the other items are checked here
    const r = G.grade(a, { workbook: back, answers: G.perfectAnswers(a) });
    const missed = r.items.filter((i) => WRITABLE.has(i.type || (a.questions.find((q) => q.id === i.id) || {}).type) && i.earned < i.points).map((i) => i.id + ': ' + i.feedback);
    assert.deepEqual(missed, [], label);
  }
});

test('an Excel upload with Power Query, a PivotTable and a chart is graded automatically', async () => {
  // Week 11 practice: a query built in Excel's Power Query and loaded to a sheet.
  const a = findAssessment('p11a');
  const key = G.buildKey(a);
  const rows = key.getSheet('Jan') && key.usedBounds('Jan').rows;
  key.queries = []; // the site's own query model is not written to .xlsx
  const jan = M_SALES.split('shared #"Store Totals"')[0];
  const bytes = await addQueries(X.write(key), key.sheets.map((s) => s.name), jan, [{ query: 'Jan', sheet: 'Jan', ref: 'A1:' + F.toAddress(rows - 1, key.usedBounds('Jan').cols - 1) }]);
  const upload = await X.read(bytes);
  const r = G.grade(a, { workbook: upload, answers: G.perfectAnswers(a) });
  assert.equal(r.percent, 100, JSON.stringify(r.items.filter((i) => i.earned < i.points)));

  // The same upload with a query that skipped steps loses the right credit.
  const lazy = await addQueries(X.write(key), key.sheets.map((s) => s.name), 'section Section1;\nshared Jan = let\n Source = Csv.Document(File.Contents("sales_jan.csv"))\nin\n Source;', [{ query: 'Jan', sheet: 'Jan', ref: 'A1:B3' }]);
  const r2 = G.grade(a, { workbook: await X.read(lazy), answers: G.perfectAnswers(a) });
  const q1 = r2.items.find((i) => i.id === 'q1');
  assert.ok(q1.earned < q1.points);
  assert.match(q1.feedback, /Remove Blank Rows|Removed Blank Rows|step/);

  // A pivot + chart question graded from Excel parts.
  const pivotQ = { id: 'px', title: 'Pivot check', workbook: { sheets: [{ name: 'Sales', rows: [['Date', 'Region', 'Revenue'], ['2026-01-05', 'East', 10], ['2026-02-05', 'West', 20]] }] },
    questions: [{ id: 'p', type: 'pivot', sheet: 'Sales', solution: { sourceRange: 'A1:C3', rows: 'Date', values: 'Revenue', agg: 'sum', rowsGroup: 'month' }, points: 2 }] };
  const book = G.buildTemplate(pivotQ);
  book.addSheet('Pivot', { silent: true });
  const withPivot = await addPivot(X.write(book), ['Sales', 'Pivot'], { sheet: 'Pivot', location: 'A3:B6', source: { sheet: 'Sales', ref: 'A1:C3' }, fields: ['Date', 'Region', 'Revenue'], rows: [0], data: { fld: 2 }, groups: [{ name: 'Date', base: 0, groupBy: 'months' }] });
  assert.equal(G.grade(pivotQ, { workbook: await X.read(withPivot) }).percent, 100);
  const noGroup = await addPivot(X.write(book), ['Sales', 'Pivot'], { sheet: 'Pivot', location: 'A3:B6', source: { sheet: 'Sales', ref: 'A1:C3' }, fields: ['Date', 'Region', 'Revenue'], rows: [0], data: { fld: 2 } });
  assert.equal(G.grade(pivotQ, { workbook: await X.read(noGroup) }).percent, 75);
});
