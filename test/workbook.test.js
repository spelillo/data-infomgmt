'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { F, W, wb } = require('./helpers');

const v = (book, sheet, addr) => { const x = book.getValue(sheet, addr); return F.isError(x) ? x.code : x; };

test('typed input is interpreted the way Excel does', () => {
  assert.deepEqual(W.interpretInput('42'), { kind: 'value', value: 42, autoFormat: W.interpretInput('42').autoFormat });
  assert.equal(W.interpretInput('$1,200.50').value, 1200.5);
  assert.equal(W.interpretInput('$1,200.50').autoFormat, 'currency');
  assert.equal(W.interpretInput('25%').value, 0.25);
  assert.equal(W.interpretInput('TRUE').value, true);
  assert.equal(W.interpretInput("'00123").value, '00123');
  assert.equal(W.interpretInput('=A1').kind, 'formula');
  assert.equal(W.interpretInput('2026-01-05').value, F.dateToSerial(2026, 1, 5));
  // inputForValue round-trips text that would otherwise be read as a number
  assert.equal(W.inputForValue('00123'), "'00123");
  assert.equal(W.inputForValue('East'), 'East');
});

test('formulas recalculate when their inputs change', () => {
  const book = wb({ S: [[1, 2, '=A1+B1']] });
  assert.equal(v(book, 'S', 'C1'), 3);
  book.setInput('S', 'A1', 10);
  assert.equal(v(book, 'S', 'C1'), 12);
});

test('cross-sheet references and sheet renames', () => {
  const book = wb({ Data: [[5]], Calc: [["=Data!A1*2"]] });
  assert.equal(v(book, 'Calc', 'A1'), 10);
  book.addSheet('My Data', { silent: true });
  book.setInput('My Data', 'A1', 7);
  book.setInput('Calc', 'B1', "='My Data'!A1+1");
  assert.equal(v(book, 'Calc', 'B1'), 8);
});

test('undo and redo restore earlier states', () => {
  const book = wb({ S: [[1]] });
  book.setInput('S', 'A1', 2);
  book.setInput('S', 'A1', 3);
  book.undo();
  assert.equal(v(book, 'S', 'A1'), 2);
  book.undo();
  assert.equal(v(book, 'S', 'A1'), 1);
  book.redo();
  assert.equal(v(book, 'S', 'A1'), 2);
});

test('locked cells refuse edits unless forced', () => {
  const book = W.Workbook.fromSpec({ sheets: [{ name: 'S', cells: { A1: { input: 'Header', locked: true } } }] });
  assert.equal(book.setInput('S', 'A1', 'changed'), false);
  assert.equal(v(book, 'S', 'A1'), 'Header');
  book.setInput('S', 'A1', 'forced', { force: true });
  assert.equal(v(book, 'S', 'A1'), 'forced');
});

test('fill down copies formulas with relative references', () => {
  const book = wb({ S: [['Units', 'Price', 'Total'], [2, 3, '=A2*B2'], [4, 5], [6, 7]] });
  book.fill('S', 'C2:C4', 'down');
  assert.equal(book.getInput('S', 'C4'), '=A4*B4');
  assert.equal(v(book, 'S', 'C4'), 42);
});

test('sortRange keeps rows together and blanks last', () => {
  const book = wb({ S: [['Name', 'Score'], ['Bo', 70], ['Al', 95], ['Cy', null], ['Di', 88]] });
  book.sortRange('S', 'A1:B5', 1, false);
  assert.deepEqual(book.getRangeValues('S', W.normRange('A2:B5')), [['Al', 95], ['Di', 88], ['Bo', 70], ['Cy', null]]);
});

test('removeDuplicates compares chosen columns, ignores case, and shifts formulas', () => {
  const book = wb({ S: [['ID', 'Name', 'Check'], ['S1', 'Ann', '=B2&"!"'], ['s1', 'ann', '=B3&"!"'], ['S2', 'Ben', '=B4&"!"']] });
  const res = book.removeDuplicates('S', 'A1:C4', true, [0, 1]);
  assert.deepEqual(res, { removed: 1, remaining: 2 });
  assert.equal(v(book, 'S', 'A3'), 'S2');
  assert.equal(book.getInput('S', 'C3'), '=B3&"!"');
  assert.equal(v(book, 'S', 'C3'), 'Ben!');
  assert.equal(v(book, 'S', 'A4'), null);
});

test('textToColumns splits and can keep leading zeros as text', () => {
  const book = wb({ S: [['Kim,0042,9'], ['Lee,0107,10']] });
  book.textToColumns('S', 'A1:A2', ',', { textColumns: [1] });
  assert.deepEqual(book.getRangeValues('S', W.normRange('A1:C2')), [['Kim', '0042', 9], ['Lee', '0107', 10]]);
});

test('importDelimited parses quoted CSV into a new sheet', () => {
  const book = wb({ S: [[1]] });
  const sheet = book.importDelimited('Name,Note\n"Smith, J","said ""hi"""\n', { sheetName: 'CSV' });
  assert.equal(sheet.name, 'CSV');
  assert.equal(v(book, 'CSV', 'A2'), 'Smith, J');
  assert.equal(v(book, 'CSV', 'B2'), 'said "hi"');
  assert.equal(W.detectDelimiter('a\tb\tc\n1\t2\t3'), '\t');
});

test('Excel Tables: names, structured references and growth', () => {
  const book = wb({ S: [['Item', 'Qty', 'Price'], ['Pens', 2, 1.5], ['Ink', 1, 9]] });
  assert.deepEqual(book.addTable('S', 'A1:C3', 'Orders'), { name: 'Orders' });
  assert.ok(book.addTable('S', 'B2:C3', 'Overlap').error);
  assert.ok(book.addTable('S', 'E1:F3', 'Bad Name').error);
  book.setInput('S', 'E1', '=SUM(Orders[Qty])');
  assert.equal(v(book, 'S', 'E1'), 3);
  book.setInput('S', 'D1', 'Total');
  book.setInput('S', 'D2', '=[@Qty]*[@Price]');
  assert.equal(book.tables[0].ref, 'A1:D3');
  assert.equal(v(book, 'S', 'D2'), 3);
  book.setInput('S', 'A4', 'Paper');
  assert.equal(book.tables[0].ref, 'A1:D4');
});

test('data validation accepts and rejects entries', () => {
  const book = wb({ S: [['Grade']] });
  book.addValidation('S', 'A2:A20', { type: 'list', values: ['9', '10', '11', '12'] });
  book.addValidation('S', 'B2:B20', { type: 'whole', min: 0, max: 100 });
  assert.equal(book.validate('S', 'A2', '11'), null);
  assert.match(book.validate('S', 'A2', '13'), /Choose one of/);
  assert.equal(book.validate('S', 'B2', '88'), null);
  assert.match(book.validate('S', 'B2', '101'), /between 0 and 100/);
  assert.match(book.validate('S', 'B2', '8.5'), /whole number/);
});

test('conditional formatting rules match the right cells', () => {
  const nums = [5, 50, 500, 50];
  const all = () => nums;
  assert.equal(W.condMatches({ op: '>', value: 100 }, 500, all), true);
  assert.equal(W.condMatches({ op: 'between', value: 10, value2: 60 }, 50, all), true);
  assert.equal(W.condMatches({ op: 'duplicate' }, 50, all), true);
  assert.equal(W.condMatches({ op: 'duplicate' }, 5, all), false);
  assert.equal(W.condMatches({ op: 'top', value: 1 }, 500, all), true);
  assert.equal(W.condMatches({ op: 'aboveAverage' }, 500, all), true);
  assert.equal(W.condMatches({ op: 'contains', value: 'east' }, 'North East', all), true);
});

test('PivotTables write a summary and refresh from the source', () => {
  const book = wb({ Sales: [['Region', 'Revenue'], ['East', 10], ['West', 5], ['East', 7]] });
  const p = book.addPivot({ sourceSheet: 'Sales', sourceRange: 'A1:B4', rows: 'Region', values: 'Revenue', agg: 'sum' });
  assert.equal(p.targetSheet, 'Pivot1');
  const read = () => book.getRangeValues('Pivot1', W.normRange('A1:B4'));
  assert.deepEqual(read(), [['Region', 'Sum of Revenue'], ['East', 17], ['West', 5], ['Grand Total', 22]]);
  book.setInput('Sales', 'B3', 50);
  book.refreshPivots();
  assert.deepEqual(read()[2], ['West', 50]);
});

test('charts read categories and series from their range', () => {
  const book = wb({ S: [['Month', 'Revenue'], ['Jan', 10], ['Feb', 20]] });
  const c = book.addChart({ sheet: 'S', type: 'column', range: 'A1:B3', title: 'Revenue' });
  const d = book.chartData(c);
  assert.deepEqual(d.categories, ['Jan', 'Feb']);
  assert.deepEqual(d.series, [{ name: 'Revenue', values: [10, 20], x: null }]);
});

test('toJSON/fromJSON preserves everything a student built', () => {
  const book = wb({ S: [['Item', 'Qty'], ['Pens', 2]] });
  book.addTable('S', 'A1:B2', 'T1');
  book.addValidation('S', 'B2:B9', { type: 'whole', min: 0, max: 9 });
  book.addChart({ sheet: 'S', type: 'bar', range: 'A1:B2' });
  const copy = W.Workbook.fromJSON(JSON.parse(JSON.stringify(book.toJSON())));
  assert.equal(copy.tables[0].name, 'T1');
  assert.equal(copy.getSheet('S').validations.length, 1);
  assert.equal(copy.charts[0].type, 'bar');
  assert.equal(v(copy, 'S', 'B2'), 2);
});

test('CSV helpers quote only when needed', () => {
  assert.equal(W.toDelimited([['a', 'b,c'], [1, 'say "x"']], ','), 'a,"b,c"\r\n1,"say ""x"""');
  assert.deepEqual(W.parseDelimited('a,"b\nc"\n1,2', ','), [['a', 'b\nc'], ['1', '2']]);
});
