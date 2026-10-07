/*
 * pivot.js — PivotTable summaries over a list of records.
 *
 * build(records, { rows, columns?, values, agg, rowsGroup? }) returns
 *   { matrix, headerRows, totalRow, rowKeys, colKeys }
 * where `matrix` is a 2D array ready to be written into a sheet:
 *
 *   Category   | Sum of Revenue        (no column field)
 *   Drinks     | 120
 *   Snacks     | 340
 *   Grand Total| 460
 *
 *   Category   | Jan | Feb | Grand Total    (with a column field)
 */
(function (root) {
  'use strict';

  const isNode = typeof module === 'object' && module.exports;
  const F = isNode ? require('./formula.js') : root.DIM.Formula;

  const AGGS = {
    sum: { label: 'Sum', fn: (xs) => F.round15(xs.reduce((a, b) => a + b, 0)), numeric: true },
    count: { label: 'Count', fn: (xs) => xs.length, numeric: false },
    average: { label: 'Average', fn: (xs) => (xs.length ? F.round15(xs.reduce((a, b) => a + b, 0) / xs.length) : null), numeric: true },
    min: { label: 'Min', fn: (xs) => (xs.length ? Math.min(...xs) : null), numeric: true },
    max: { label: 'Max', fn: (xs) => (xs.length ? Math.max(...xs) : null), numeric: true },
  };

  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // Turn a raw field value into a group key and a sortable value.
  function groupKey(v, mode) {
    if (v === null || v === undefined || v === '') return { label: '(blank)', sort: [3, ''] };
    if (mode && typeof v === 'number') {
      const d = F.serialToDate(v);
      if (mode === 'year') return { label: String(d.y), sort: [0, d.y] };
      if (mode === 'quarter') { const q = Math.ceil(d.m / 3); return { label: 'Q' + q + ' ' + d.y, sort: [0, d.y * 10 + q] }; }
      if (mode === 'month') return { label: MONTHS[d.m - 1] + ' ' + d.y, sort: [0, d.y * 100 + d.m] };
    }
    if (typeof v === 'number') return { label: v, sort: [0, v] };
    if (typeof v === 'boolean') return { label: v ? 'TRUE' : 'FALSE', sort: [2, v ? 1 : 0] };
    return { label: String(v), sort: [1, String(v).toLowerCase()] };
  }

  function cmpSort(a, b) {
    if (a[0] !== b[0]) return a[0] - b[0];
    return a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0;
  }

  function findField(record, name) {
    if (!name) return undefined;
    if (Object.prototype.hasOwnProperty.call(record, name)) return name;
    const lower = String(name).toLowerCase();
    return Object.keys(record).find((k) => k.toLowerCase() === lower);
  }

  function build(records, opts) {
    const agg = AGGS[opts.agg || 'sum'];
    if (!agg) return { error: 'Unknown summary "' + opts.agg + '".' };
    if (!records.length) return { error: 'The source has no rows.' };
    const rowField = findField(records[0], opts.rows);
    const colField = opts.columns ? findField(records[0], opts.columns) : null;
    const valField = findField(records[0], opts.values);
    if (!rowField) return { error: 'Choose a field for Rows.' };
    if (opts.columns && !colField) return { error: 'Column field "' + opts.columns + '" was not found.' };
    if (!valField) return { error: 'Choose a field for Values.' };

    const rowKeys = new Map();
    const colKeys = new Map();
    const cells = new Map(); // rowLabel \u0001 colLabel -> [values]
    const push = (k, v) => { if (!cells.has(k)) cells.set(k, []); cells.get(k).push(v); };

    for (const rec of records) {
      const rk = groupKey(rec[rowField], opts.rowsGroup);
      rowKeys.set(String(rk.label), rk);
      const ck = colField ? groupKey(rec[colField], opts.columnsGroup) : { label: '', sort: [0, 0] };
      colKeys.set(String(ck.label), ck);
      const v = rec[valField];
      if (agg.numeric) {
        if (typeof v !== 'number') continue;
      } else if (v === null || v === undefined || v === '') continue;
      const val = agg.numeric ? v : 1;
      push(rk.label + '\u0001' + ck.label, val);
      push(rk.label + '\u0001*', val);
      push('*\u0001' + ck.label, val);
      push('*\u0001*', val);
    }

    const rows = [...rowKeys.values()].sort((a, b) => cmpSort(a.sort, b.sort));
    const cols = [...colKeys.values()].sort((a, b) => cmpSort(a.sort, b.sort));
    const get = (r, c) => {
      const xs = cells.get(r + '\u0001' + c);
      return xs && xs.length ? agg.fn(xs) : null;
    };
    const valueLabel = agg.label + ' of ' + valField;
    const matrix = [];
    if (!colField) {
      matrix.push([rowField, valueLabel]);
      rows.forEach((r) => matrix.push([r.label, get(r.label, '')]));
      matrix.push(['Grand Total', get('*', '*')]);
      return { matrix, headerRows: 1, totalRow: true, rowKeys: rows.map((r) => r.label), colKeys: [], valueLabel };
    }
    matrix.push([valueLabel, colField + ' →']);
    matrix.push([rowField, ...cols.map((c) => c.label), 'Grand Total']);
    rows.forEach((r) => matrix.push([r.label, ...cols.map((c) => get(r.label, c.label)), get(r.label, '*')]));
    matrix.push(['Grand Total', ...cols.map((c) => get('*', c.label)), get('*', '*')]);
    return { matrix, headerRows: 2, totalRow: true, rowKeys: rows.map((r) => r.label), colKeys: cols.map((c) => c.label), valueLabel };
  }

  const api = { build, AGGS };
  root.DIM = root.DIM || {};
  root.DIM.Pivot = api;
  if (isNode) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
