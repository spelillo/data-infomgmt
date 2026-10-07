'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { Q, W, C, wb } = require('./helpers');

const CSV = [
  'Order ID,Store,Product ID,Qty,Unit Price,Order Date',
  '1001, cafeteria ,P1,2,1.50,2026-01-05',
  '1002,Library,P2,1,3.00,2026-01-06',
  ',,,,,',
  '1002,Library,P2,1,3.00,2026-01-06',
  '1003,Gym,P9,x,2.00,2026-01-07',
].join('\n');
const PRODUCTS = 'ProductID,Category\nP1,Supplies\nP2,Books\n';

function book() {
  const b = wb({ Data: [['Club', 'Jan', 'Feb'], ['Art', 10, 12], ['Chess', 7, null]] });
  b.sources['orders.csv'] = CSV;
  b.sources['more.csv'] = 'Order ID,Store,Product ID,Qty,Unit Price,Order Date\n1004,Gym,P1,3,1.50,2026-02-01\n';
  b.sources['products.csv'] = PRODUCTS;
  return b;
}
const run = (steps, source = { kind: 'file', file: 'orders.csv' }, b = book()) => Q.run({ name: 'T', source, steps }, b);
const col = (res, c) => res.table.rows.map((r) => r[c]);

test('a file source loads every column as text, like Power Query before types', () => {
  const r = run([]);
  assert.deepEqual(r.table.columns, ['Order ID', 'Store', 'Product ID', 'Qty', 'Unit Price', 'Order Date']);
  assert.equal(r.table.rows.length, 4); // the all-blank CSV line is dropped by the reader
  assert.equal(col(r, 'Qty')[0], '2');
});

test('cleaning steps: blank rows, duplicates, trim, proper', () => {
  const r = run([
    { type: 'removeBlankRows' },
    { type: 'removeDuplicates', columns: [] },
    { type: 'transform', column: 'Store', op: 'trim' },
    { type: 'transform', column: 'Store', op: 'proper' },
  ]);
  assert.deepEqual(col(r, 'Store'), ['Cafeteria', 'Library', 'Gym']);
});

test('changeType converts values and turns bad ones into errors that can be removed', () => {
  const typed = run([{ type: 'changeType', columns: { Qty: 'whole', 'Unit Price': 'currency', 'Order Date': 'date' } }]);
  assert.equal(col(typed, 'Qty')[0], 2);
  assert.ok(typed.errorCells >= 1);
  const cleaned = run([{ type: 'changeType', columns: { Qty: 'whole' } }, { type: 'removeErrors' }]);
  assert.equal(cleaned.errorCells, 0);
  assert.ok(!col(cleaned, 'Order ID').includes('1003'));
});

test('custom columns, filters, sorting and keeping top rows', () => {
  const r = run([
    { type: 'removeBlankRows' },
    { type: 'changeType', columns: { Qty: 'whole', 'Unit Price': 'number' } },
    { type: 'removeErrors' },
    { type: 'addColumn', name: 'Revenue', formula: '[Qty]*[Unit Price]' },
    { type: 'filterRows', column: 'Revenue', op: '>', value: 2 },
    { type: 'sortRows', column: 'Revenue', order: 'desc' },
    { type: 'keepTopRows', count: 1 },
  ]);
  assert.deepEqual(col(r, 'Revenue'), [3]);
});

test('column steps: rename, remove, keep, split', () => {
  const r = run([
    { type: 'renameColumn', from: 'Order Date', to: 'Date' },
    { type: 'keepColumns', columns: ['Order ID', 'Date'] },
    { type: 'splitColumn', column: 'Date', delimiter: '-', names: ['Y', 'M', 'D'] },
    { type: 'removeColumns', columns: ['D'] },
  ]);
  assert.deepEqual(r.table.columns, ['Order ID', 'Y', 'M']);
  const bad = run([{ type: 'renameColumn', from: 'Nope', to: 'X' }]);
  assert.match(bad.error, /Nope/);
  assert.equal(bad.errorStep, 0);
});

test('group by summarizes rows', () => {
  const r = run([
    { type: 'removeBlankRows' },
    { type: 'changeType', columns: { Qty: 'whole' } },
    { type: 'removeErrors' },
    { type: 'groupBy', by: ['Store'], aggregations: [{ name: 'Units', op: 'sum', column: 'Qty' }, { name: 'Orders', op: 'count' }] },
  ]);
  assert.deepEqual(r.table.rows.find((x) => x.Store === 'Library'), { Store: 'Library', Units: 2, Orders: 2 });
});

test('unpivot turns a wide sheet long and drops empty cells', () => {
  const r = run([{ type: 'unpivot', keep: ['Club'] }], { kind: 'sheet', sheet: 'Data' });
  assert.deepEqual(r.table.columns, ['Club', 'Attribute', 'Value']);
  assert.equal(r.table.rows.length, 3);
});

test('merge looks up columns from another source; left anti finds orphans', () => {
  const steps = (kind) => [{ type: 'removeBlankRows' }, { type: 'merge', with: { kind: 'file', file: 'products.csv' }, leftKey: 'Product ID', rightKey: 'ProductID', kind, columns: ['Category'] }];
  const left = run(steps('leftOuter'));
  assert.deepEqual(col(left, 'Category'), ['Supplies', 'Books', 'Books', null]);
  const anti = run(steps('leftAnti'));
  assert.deepEqual(col(anti, 'Product ID'), ['P9']);
});

test('folder sources append files and add Source.Name', () => {
  const r = run([], { kind: 'folder', files: ['orders.csv', 'more.csv'] });
  assert.equal(r.table.columns[0], 'Source.Name');
  assert.equal(r.table.rows.length, 5);
  assert.equal(r.table.rows.at(-1)['Source.Name'], 'more.csv');
});

test('Close & Load writes a sheet and Refresh picks up new data', () => {
  const b = book();
  b.saveQuery({ id: 'q1', name: 'Orders', source: { kind: 'file', file: 'orders.csv' }, steps: [{ type: 'removeBlankRows' }, { type: 'removeDuplicates', columns: [] }] });
  b.loadQuery('q1');
  assert.equal(b.getValue('Orders', 'A1'), 'Order ID');
  assert.equal(b.queries[0].loadedRows, 3);
  b.sources['orders.csv'] += '\n1005,Gym,P1,1,1.50,2026-01-09';
  b.refreshAll();
  assert.equal(b.queries[0].loadedRows, 4);
});

test('every step type has a label and a description', () => {
  for (const type of Object.keys(Q.STEPS)) {
    assert.ok(Q.STEP_LABELS[type], 'label for ' + type);
    assert.equal(typeof Q.describe({ type }), 'string');
  }
});

test('course data files load cleanly', () => {
  for (const [name, text] of Object.entries(C.data.files)) {
    const t = Q.tableFromText(text);
    assert.ok(t.columns.length > 1, name + ' has columns');
    assert.ok(t.rows.length > 0, name + ' has rows');
  }
});
