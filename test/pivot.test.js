'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { F, Pivot } = require('./helpers');

const d = (y, m, day) => F.dateToSerial(y, m, day);
const RECORDS = [
  { Region: 'East', Category: 'Pens', Revenue: 10, Date: d(2026, 1, 5) },
  { Region: 'West', Category: 'Ink', Revenue: 30, Date: d(2026, 1, 20) },
  { Region: 'East', Category: 'Ink', Revenue: 20, Date: d(2026, 2, 3) },
  { Region: 'North', Category: 'Pens', Revenue: 5, Date: d(2026, 4, 9) },
  { Region: 'East', Category: 'Pens', Revenue: 'n/a', Date: d(2026, 4, 30) },
];

test('rows + values with sum, sorted labels and a grand total', () => {
  const out = Pivot.build(RECORDS, { rows: 'Region', values: 'Revenue', agg: 'sum' });
  assert.deepEqual(out.matrix, [['Region', 'Sum of Revenue'], ['East', 30], ['North', 5], ['West', 30], ['Grand Total', 65]]);
});

test('count counts non-blank entries, including text', () => {
  const out = Pivot.build(RECORDS, { rows: 'Region', values: 'Revenue', agg: 'count' });
  assert.deepEqual(out.matrix[1], ['East', 3]);
  assert.deepEqual(out.matrix.at(-1), ['Grand Total', 5]);
});

test('average, min and max ignore text values', () => {
  assert.deepEqual(Pivot.build(RECORDS, { rows: 'Region', values: 'Revenue', agg: 'average' }).matrix[1], ['East', 15]);
  assert.deepEqual(Pivot.build(RECORDS, { rows: 'Category', values: 'Revenue', agg: 'max' }).matrix[1], ['Ink', 30]);
  assert.deepEqual(Pivot.build(RECORDS, { rows: 'Category', values: 'Revenue', agg: 'min' }).matrix[2], ['Pens', 5]);
});

test('a column field makes a cross-tab', () => {
  const out = Pivot.build(RECORDS, { rows: 'Region', columns: 'Category', values: 'Revenue', agg: 'sum' });
  assert.deepEqual(out.matrix[1], ['Region', 'Ink', 'Pens', 'Grand Total']);
  assert.deepEqual(out.matrix[2], ['East', 20, 10, 30]);
  assert.deepEqual(out.matrix.at(-1), ['Grand Total', 50, 15, 65]);
});

test('dates group by month, quarter or year in calendar order', () => {
  const m = Pivot.build(RECORDS, { rows: 'Date', values: 'Revenue', agg: 'sum', rowsGroup: 'month' });
  assert.deepEqual(m.matrix.slice(1, -1), [['Jan 2026', 40], ['Feb 2026', 20], ['Apr 2026', 5]]);
  const q = Pivot.build(RECORDS, { rows: 'Date', values: 'Revenue', agg: 'sum', rowsGroup: 'quarter' });
  assert.deepEqual(q.matrix.slice(1, -1), [['Q1 2026', 60], ['Q2 2026', 5]]);
  const y = Pivot.build(RECORDS, { rows: 'Date', values: 'Revenue', agg: 'sum', rowsGroup: 'year' });
  assert.deepEqual(y.matrix.slice(1, -1), [['2026', 65]]);
});

test('field names match without regard to case, and errors explain themselves', () => {
  assert.equal(Pivot.build(RECORDS, { rows: 'region', values: 'REVENUE' }).matrix[1][1], 30);
  assert.match(Pivot.build(RECORDS, { rows: 'Nope', values: 'Revenue' }).error, /Rows/);
  assert.match(Pivot.build(RECORDS, { rows: 'Region', values: 'Revenue', agg: 'mode' }).error, /Unknown summary/);
  assert.match(Pivot.build([], { rows: 'Region', values: 'Revenue' }).error, /no rows/);
});
