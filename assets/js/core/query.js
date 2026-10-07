/*
 * query.js — a teaching-sized Power Query.
 *
 * A query is a source plus an ordered list of "Applied Steps". Every step is a
 * pure function table -> table, so a query can be re-run ("Refresh") against
 * new data and get the same cleaning applied again. That repeatability is the
 * whole point of Power Query, and it is what Unit 3 teaches.
 *
 * Table shape: { columns: ['Region', 'Sales'], rows: [{ Region: 'East', Sales: 10 }], types: { Sales: 'number' } }
 *
 * Query definition:
 *   { id, name, source: { kind: 'file', file: 'sales_jan.csv' }
 *                     | { kind: 'folder', files: ['jan.csv', 'feb.csv'] }
 *                     | { kind: 'sheet', sheet: 'Data', range?: 'A1:D20' }
 *                     | { kind: 'query', name: 'Products' },
 *     steps: [ { type: 'changeType', columns: { Sales: 'number' } }, ... ] }
 */
(function (root) {
  'use strict';

  const isNode = typeof module === 'object' && module.exports;
  const F = isNode ? require('./formula.js') : root.DIM.Formula;
  const W = () => (isNode ? require('./workbook.js') : root.DIM.Workbook);

  const ERROR = F.err('#VALUE!', 'Power Query could not convert this value');
  const isErr = (v) => F.isError(v);
  const blank = (v) => v === null || v === undefined || v === '';

  // Human-readable labels for the Applied Steps pane.
  const STEP_LABELS = {
    changeType: 'Changed Type',
    removeColumns: 'Removed Columns',
    keepColumns: 'Removed Other Columns',
    renameColumn: 'Renamed Column',
    filterRows: 'Filtered Rows',
    replaceValues: 'Replaced Value',
    transform: 'Transformed Text',
    splitColumn: 'Split Column by Delimiter',
    removeDuplicates: 'Removed Duplicates',
    removeBlankRows: 'Removed Blank Rows',
    removeErrors: 'Removed Errors',
    sortRows: 'Sorted Rows',
    addColumn: 'Added Custom Column',
    groupBy: 'Grouped Rows',
    unpivot: 'Unpivoted Columns',
    merge: 'Merged Queries',
    append: 'Appended Query',
    keepTopRows: 'Kept First Rows',
    fillDown: 'Filled Down',
  };

  function describe(step) {
    const base = STEP_LABELS[step.type] || step.type;
    switch (step.type) {
      case 'changeType': return base + ': ' + Object.entries(step.columns || {}).map(([c, t]) => c + ' → ' + t).join(', ');
      case 'removeColumns': return base + ': ' + (step.columns || []).join(', ');
      case 'keepColumns': return 'Kept Columns: ' + (step.columns || []).join(', ');
      case 'renameColumn': return base + ': ' + step.from + ' → ' + step.to;
      case 'filterRows': return base + ': ' + step.column + ' ' + step.op + (step.value !== undefined && step.value !== '' ? ' ' + step.value : '');
      case 'replaceValues': return base + ' in ' + step.column + ': "' + step.find + '" → "' + step.replace + '"';
      case 'transform': return base + ' (' + step.op + '): ' + step.column;
      case 'splitColumn': return base + ' "' + step.delimiter + '": ' + step.column;
      case 'sortRows': return base + ': ' + step.column + ' ' + (step.order === 'desc' ? 'Z→A' : 'A→Z');
      case 'addColumn': return base + ': ' + step.name + ' = ' + step.formula;
      case 'groupBy': return base + ' by ' + [].concat(step.by).join(', ');
      case 'unpivot': return 'Unpivoted Other Columns (kept ' + (step.keep || []).join(', ') + ')';
      case 'merge': return base + ' with ' + sourceLabel(step.with) + ' (' + (step.kind || 'leftOuter') + ')';
      case 'append': return base + ': ' + [].concat(step.with).map(sourceLabel).join(', ');
      case 'keepTopRows': return base + ': ' + step.count;
      case 'fillDown': return base + ': ' + step.column;
      case 'removeDuplicates': return base + (step.columns && step.columns.length ? ': ' + step.columns.join(', ') : '');
      default: return base;
    }
  }

  function sourceLabel(src) {
    if (!src) return '?';
    if (typeof src === 'string') return src;
    if (src.kind === 'file') return src.file;
    if (src.kind === 'folder') return 'Folder (' + src.files.length + ' files)';
    if (src.kind === 'sheet') return 'Sheet ' + src.sheet;
    if (src.kind === 'query') return src.name;
    return '?';
  }

  // ---------------------------------------------------------------------------
  // Sources
  // ---------------------------------------------------------------------------

  function tableFromText(text, delimiter) {
    const Wb = W();
    const delim = delimiter === 'tab' ? '\t' : delimiter || Wb.detectDelimiter(text);
    const rows = Wb.parseDelimited(text, delim).filter((r) => r.some((x) => x !== ''));
    if (!rows.length) return { columns: [], rows: [], types: {} };
    const header = uniqueNames(rows[0].map((h, i) => (h.trim() === '' ? 'Column' + (i + 1) : h.trim())));
    const out = rows.slice(1).map((r) => {
      const o = {};
      header.forEach((h, i) => { o[h] = r[i] === undefined || r[i] === '' ? null : r[i]; });
      return o;
    });
    return { columns: header, rows: out, types: Object.fromEntries(header.map((h) => [h, 'text'])) };
  }

  function uniqueNames(names) {
    const seen = {};
    return names.map((n) => {
      if (!seen[n]) { seen[n] = 1; return n; }
      seen[n]++;
      return n + '.' + seen[n];
    });
  }

  function loadSource(src, wb, depth = 0) {
    if (depth > 8) return { error: 'Queries refer to each other in a loop.' };
    if (!src) return { error: 'This query has no source.' };
    if (src.kind === 'file') {
      const text = wb.sources[src.file];
      if (text === undefined) return { error: 'File "' + src.file + '" was not found.' };
      return { table: tableFromText(text, src.delimiter) };
    }
    if (src.kind === 'folder') {
      const parts = [];
      for (const f of src.files || []) {
        const r = loadSource({ kind: 'file', file: f, delimiter: src.delimiter }, wb, depth + 1);
        if (r.error) return r;
        if (src.addSourceColumn !== false) {
          r.table.columns = ['Source.Name', ...r.table.columns];
          r.table.rows.forEach((row) => { row['Source.Name'] = f; });
          r.table.types['Source.Name'] = 'text';
        }
        parts.push(r.table);
      }
      return { table: appendTables(parts) };
    }
    if (src.kind === 'sheet') {
      const rec = wb.recordsFrom(src.sheet, src.range);
      if (rec.error) return rec;
      const types = {};
      rec.columns.forEach((c) => { types[c] = 'any'; });
      return { table: { columns: rec.columns, rows: rec.rows, types } };
    }
    if (src.kind === 'query') {
      const def = wb.queries.find((q) => q.name.toLowerCase() === String(src.name).toLowerCase());
      if (!def) return { error: 'Query "' + src.name + '" was not found.' };
      const r = run(def, wb, undefined, depth + 1);
      return r.error ? r : { table: r.table };
    }
    return { error: 'Unknown source.' };
  }

  // ---------------------------------------------------------------------------
  // Value conversion
  // ---------------------------------------------------------------------------

  function convert(v, type) {
    if (blank(v)) return null;
    if (isErr(v)) return v;
    switch (type) {
      case 'text': return typeof v === 'number' ? F.formatGeneral(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
      case 'number':
      case 'currency': {
        if (typeof v === 'number') return v;
        const p = F.parseNumberLike(String(v).trim());
        return p && p.format !== 'date' ? p.value : ERROR;
      }
      case 'whole': {
        const n = convert(v, 'number');
        return isErr(n) ? n : Math.round(n);
      }
      case 'date': {
        if (typeof v === 'number') return Math.floor(v);
        const p = F.parseNumberLike(String(v).trim());
        return p && p.format === 'date' ? p.value : ERROR;
      }
      case 'boolean': {
        const s = String(v).trim().toLowerCase();
        if (['true', 'yes', '1'].includes(s)) return true;
        if (['false', 'no', '0'].includes(s)) return false;
        return ERROR;
      }
      default: return v;
    }
  }

  function textOf(v) { return blank(v) ? '' : isErr(v) ? '' : typeof v === 'string' ? v : F.toText(v); }

  // ---------------------------------------------------------------------------
  // Steps
  // ---------------------------------------------------------------------------

  function cloneTable(t) {
    return { columns: [...t.columns], rows: t.rows.map((r) => ({ ...r })), types: { ...(t.types || {}) } };
  }

  function needColumn(t, c) {
    if (!t.columns.includes(c)) throw new Error('Column "' + c + '" was not found. Columns are: ' + t.columns.join(', '));
  }

  function appendTables(tables) {
    const columns = [];
    tables.forEach((t) => t.columns.forEach((c) => { if (!columns.includes(c)) columns.push(c); }));
    const rows = [];
    tables.forEach((t) => t.rows.forEach((r) => {
      const o = {};
      columns.forEach((c) => { o[c] = r[c] === undefined ? null : r[c]; });
      rows.push(o);
    }));
    const types = {};
    columns.forEach((c) => {
      const ts = tables.map((t) => t.types && t.types[c]).filter(Boolean);
      types[c] = ts.length && ts.every((x) => x === ts[0]) ? ts[0] : 'any';
    });
    return { columns, rows, types };
  }

  function keyOf(v) {
    if (blank(v)) return '\u0000';
    return typeof v === 'string' ? 's:' + v : typeof v + ':' + String(v);
  }

  function filterTest(step) {
    const target = step.value;
    const tnum = typeof target === 'number' ? target : F.parseNumberLike(String(target ?? ''));
    const numTarget = typeof target === 'number' ? target : tnum ? tnum.value : null;
    return (v) => {
      switch (step.op) {
        case 'blank': return blank(v);
        case 'notBlank': return !blank(v);
        case 'contains': return textOf(v).toLowerCase().includes(String(target).toLowerCase());
        case 'notContains': return !textOf(v).toLowerCase().includes(String(target).toLowerCase());
        case 'beginsWith': return textOf(v).toLowerCase().startsWith(String(target).toLowerCase());
        case 'endsWith': return textOf(v).toLowerCase().endsWith(String(target).toLowerCase());
      }
      if (blank(v)) return step.op === '<>';
      if (typeof v === 'number' && numTarget !== null) {
        switch (step.op) {
          case '=': return v === numTarget;
          case '<>': return v !== numTarget;
          case '>': return v > numTarget;
          case '<': return v < numTarget;
          case '>=': return v >= numTarget;
          case '<=': return v <= numTarget;
        }
      }
      const c = F.compare(textOf(v).toLowerCase(), String(target ?? '').toLowerCase());
      switch (step.op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '>': return c > 0;
        case '<': return c < 0;
        case '>=': return c >= 0;
        case '<=': return c <= 0;
      }
      return true;
    };
  }

  function properCase(s) {
    return s.toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (m, p, ch) => p + ch.toUpperCase());
  }

  // Evaluate a custom-column formula such as  [Qty] * [Price]  or  UPPER([Name]).
  function compileCustom(formula, columns) {
    const names = [];
    const body = String(formula).replace(/^=/, '').replace(/\[([^\]]+)\]/g, (m, col) => {
      const c = columns.find((x) => x.toLowerCase() === col.trim().toLowerCase());
      if (!c) throw new Error('Column [' + col + '] was not found.');
      let i = names.indexOf(c);
      if (i < 0) { names.push(c); i = names.length - 1; }
      return '__COL' + i;
    });
    const ast = F.parse('=' + body);
    return (row) => {
      const ctx = {
        sheet: '__row', row: 0, col: 0,
        getCell: (sheet, r, c) => {
          const v = row[names[c]];
          return blank(v) ? null : v;
        },
        resolveName: (n) => {
          const m = /^__COL(\d+)$/i.exec(n);
          return m ? new F.RangeRef('__row', 0, +m[1], 0, +m[1]) : null;
        },
      };
      const v = F.deref(F.evaluate(ast, ctx), ctx);
      return v === undefined ? null : v;
    };
  }

  const STEPS = {
    changeType(t, s) {
      for (const [c, type] of Object.entries(s.columns || {})) {
        needColumn(t, c);
        t.rows.forEach((r) => { r[c] = convert(r[c], type); });
        t.types[c] = type;
      }
      return t;
    },
    removeColumns(t, s) {
      (s.columns || []).forEach((c) => needColumn(t, c));
      t.columns = t.columns.filter((c) => !s.columns.includes(c));
      t.rows.forEach((r) => s.columns.forEach((c) => delete r[c]));
      s.columns.forEach((c) => delete t.types[c]);
      return t;
    },
    keepColumns(t, s) {
      (s.columns || []).forEach((c) => needColumn(t, c));
      const drop = t.columns.filter((c) => !s.columns.includes(c));
      return STEPS.removeColumns(t, { columns: drop });
    },
    renameColumn(t, s) {
      needColumn(t, s.from);
      if (!s.to || (t.columns.includes(s.to) && s.to !== s.from)) throw new Error('Column name "' + s.to + '" is empty or already used.');
      t.columns = t.columns.map((c) => (c === s.from ? s.to : c));
      t.rows.forEach((r) => { r[s.to] = r[s.from]; if (s.to !== s.from) delete r[s.from]; });
      t.types[s.to] = t.types[s.from]; if (s.to !== s.from) delete t.types[s.from];
      return t;
    },
    filterRows(t, s) {
      needColumn(t, s.column);
      const test = filterTest(s);
      t.rows = t.rows.filter((r) => !isErr(r[s.column]) && test(r[s.column]));
      return t;
    },
    replaceValues(t, s) {
      const cols = s.column ? [s.column] : t.columns;
      cols.forEach((c) => needColumn(t, c));
      const find = s.find === undefined ? '' : String(s.find);
      t.rows.forEach((r) => cols.forEach((c) => {
        const v = r[c];
        if (s.matchEntire || typeof v !== 'string') {
          if (textOf(v) === find || (find === '' && blank(v))) {
            r[c] = s.replace === '' ? null : t.types[c] && t.types[c] !== 'text' && t.types[c] !== 'any' ? convert(s.replace, t.types[c]) : s.replace;
          }
        } else if (find !== '') {
          r[c] = v.split(find).join(s.replace ?? '');
        }
      }));
      return t;
    },
    transform(t, s) {
      needColumn(t, s.column);
      const fn = {
        trim: (x) => x.replace(/\s+/g, ' ').trim(),
        clean: (x) => x.replace(/[\x00-\x1f]/g, ''),
        upper: (x) => x.toUpperCase(),
        lower: (x) => x.toLowerCase(),
        proper: properCase,
      }[s.op];
      if (!fn) throw new Error('Unknown text transform "' + s.op + '".');
      t.rows.forEach((r) => { if (typeof r[s.column] === 'string') r[s.column] = fn(r[s.column]); });
      return t;
    },
    splitColumn(t, s) {
      needColumn(t, s.column);
      const delim = s.delimiter === 'tab' ? '\t' : s.delimiter === 'space' ? ' ' : s.delimiter;
      if (!delim) throw new Error('Choose a delimiter.');
      const parts = t.rows.map((r) => (blank(r[s.column]) ? [] : textOf(r[s.column]).split(delim).map((x) => x.trim())));
      const n = Math.max(1, ...parts.map((p) => p.length));
      const names = uniqueAgainst(Array.from({ length: n }, (_, i) => (s.names && s.names[i]) || s.column + '.' + (i + 1)), t.columns.filter((c) => c !== s.column));
      const at = t.columns.indexOf(s.column);
      t.columns.splice(at, 1, ...names);
      t.rows.forEach((r, i) => {
        delete r[s.column];
        names.forEach((nm, j) => { r[nm] = parts[i][j] === undefined || parts[i][j] === '' ? null : parts[i][j]; });
      });
      delete t.types[s.column];
      names.forEach((nm) => { t.types[nm] = 'text'; });
      return t;
    },
    removeDuplicates(t, s) {
      const cols = s.columns && s.columns.length ? s.columns : t.columns;
      cols.forEach((c) => needColumn(t, c));
      const seen = new Set();
      t.rows = t.rows.filter((r) => {
        const k = cols.map((c) => keyOf(typeof r[c] === 'string' ? r[c] : r[c])).join('\u0001');
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });
      return t;
    },
    removeBlankRows(t) {
      t.rows = t.rows.filter((r) => t.columns.some((c) => !blank(r[c])));
      return t;
    },
    removeErrors(t, s) {
      const cols = s.column ? [s.column] : t.columns;
      t.rows = t.rows.filter((r) => !cols.some((c) => isErr(r[c])));
      return t;
    },
    sortRows(t, s) {
      needColumn(t, s.column);
      const dir = s.order === 'desc' ? -1 : 1;
      t.rows = t.rows.map((r, i) => ({ r, i })).sort((a, b) => {
        const x = a.r[s.column], y = b.r[s.column];
        if (blank(x) !== blank(y)) return blank(x) ? 1 : -1;
        return dir * F.compare(isErr(x) ? null : x, isErr(y) ? null : y) || a.i - b.i;
      }).map((x) => x.r);
      return t;
    },
    addColumn(t, s) {
      if (!s.name) throw new Error('Name the new column.');
      if (t.columns.includes(s.name)) throw new Error('There is already a column named "' + s.name + '".');
      const fn = compileCustom(s.formula, t.columns);
      t.rows.forEach((r) => { r[s.name] = fn(r); });
      t.columns.push(s.name);
      t.types[s.name] = s.dataType || 'any';
      return t;
    },
    groupBy(t, s) {
      const by = [].concat(s.by || []);
      by.forEach((c) => needColumn(t, c));
      const aggs = s.aggregations || [];
      aggs.forEach((a) => { if (a.op !== 'count') needColumn(t, a.column); });
      const groups = new Map();
      t.rows.forEach((r) => {
        const k = by.map((c) => keyOf(r[c])).join('\u0001');
        if (!groups.has(k)) groups.set(k, { key: by.map((c) => r[c]), rows: [] });
        groups.get(k).rows.push(r);
      });
      const rows = [...groups.values()].map((g) => {
        const o = {};
        by.forEach((c, i) => { o[c] = g.key[i]; });
        aggs.forEach((a) => {
          const nums = a.op === 'count' ? [] : g.rows.map((r) => r[a.column]).filter((v) => typeof v === 'number');
          switch (a.op) {
            case 'count': o[a.name] = g.rows.length; break;
            case 'sum': o[a.name] = F.round15(nums.reduce((x, y) => x + y, 0)); break;
            case 'average': o[a.name] = nums.length ? F.round15(nums.reduce((x, y) => x + y, 0) / nums.length) : null; break;
            case 'min': o[a.name] = nums.length ? Math.min(...nums) : null; break;
            case 'max': o[a.name] = nums.length ? Math.max(...nums) : null; break;
            case 'countDistinct': o[a.name] = new Set(g.rows.map((r) => keyOf(r[a.column]))).size; break;
            default: throw new Error('Unknown aggregation "' + a.op + '".');
          }
        });
        return o;
      });
      const types = {};
      by.forEach((c) => { types[c] = t.types[c]; });
      aggs.forEach((a) => { types[a.name] = 'number'; });
      return { columns: [...by, ...aggs.map((a) => a.name)], rows, types };
    },
    unpivot(t, s) {
      const keep = s.keep || [];
      keep.forEach((c) => needColumn(t, c));
      const attr = s.attributeName || 'Attribute', val = s.valueName || 'Value';
      const others = t.columns.filter((c) => !keep.includes(c));
      const rows = [];
      t.rows.forEach((r) => others.forEach((c) => {
        if (blank(r[c])) return; // Power Query drops nulls when unpivoting
        const o = {};
        keep.forEach((k) => { o[k] = r[k]; });
        o[attr] = c;
        o[val] = r[c];
        rows.push(o);
      }));
      const types = {};
      keep.forEach((k) => { types[k] = t.types[k]; });
      types[attr] = 'text';
      const vt = others.map((c) => t.types[c]);
      types[val] = vt.length && vt.every((x) => x === vt[0]) ? vt[0] : 'any';
      return { columns: [...keep, attr, val], rows, types };
    },
    merge(t, s, wb, depth) {
      const other = loadOther(s.with, wb, depth);
      needColumn(t, s.leftKey);
      if (!other.columns.includes(s.rightKey)) throw new Error('Column "' + s.rightKey + '" was not found in ' + sourceLabel(s.with) + '.');
      const kind = s.kind || 'leftOuter';
      const expand = s.columns && s.columns.length ? s.columns : other.columns.filter((c) => c !== s.rightKey);
      expand.forEach((c) => { if (!other.columns.includes(c)) throw new Error('Column "' + c + '" was not found in ' + sourceLabel(s.with) + '.'); });
      const prefix = s.prefix === undefined ? '' : s.prefix;
      const newNames = uniqueAgainst(expand.map((c) => prefix + c), t.columns);
      const index = new Map();
      other.rows.forEach((r, i) => {
        const k = matchKey(r[s.rightKey]);
        if (!index.has(k)) index.set(k, []);
        index.get(k).push(i);
      });
      const usedRight = new Set();
      const rows = [];
      const emptyRight = Object.fromEntries(newNames.map((n) => [n, null]));
      t.rows.forEach((r) => {
        const hits = blank(r[s.leftKey]) ? [] : index.get(matchKey(r[s.leftKey])) || [];
        hits.forEach((i) => usedRight.add(i));
        if (kind === 'leftAnti') { if (!hits.length) rows.push({ ...r, ...emptyRight }); return; }
        if (kind === 'rightAnti') return;
        if (!hits.length) {
          if (kind === 'leftOuter' || kind === 'fullOuter') rows.push({ ...r, ...emptyRight });
          return;
        }
        hits.forEach((i) => {
          const o = { ...r };
          expand.forEach((c, j) => { o[newNames[j]] = other.rows[i][c]; });
          rows.push(o);
        });
      });
      if (kind === 'rightOuter' || kind === 'fullOuter' || kind === 'rightAnti') {
        if (kind === 'rightOuter') {
          // right outer keeps every right row; rebuild in right-table order
          rows.length = 0;
          other.rows.forEach((rr, i) => {
            const lefts = t.rows.filter((r) => !blank(r[s.leftKey]) && matchKey(r[s.leftKey]) === matchKey(rr[s.rightKey]));
            const right = Object.fromEntries(expand.map((c, j) => [newNames[j], rr[c]]));
            if (!lefts.length) rows.push({ ...Object.fromEntries(t.columns.map((c) => [c, null])), ...right });
            else lefts.forEach((l) => rows.push({ ...l, ...right }));
          });
        } else {
          other.rows.forEach((rr, i) => {
            if (usedRight.has(i)) return;
            const right = Object.fromEntries(expand.map((c, j) => [newNames[j], rr[c]]));
            rows.push({ ...Object.fromEntries(t.columns.map((c) => [c, null])), ...right });
          });
        }
      }
      const types = { ...t.types };
      expand.forEach((c, j) => { types[newNames[j]] = other.types[c] || 'any'; });
      return { columns: [...t.columns, ...newNames], rows, types };
    },
    append(t, s, wb, depth) {
      const others = [].concat(s.with).map((w) => loadOther(w, wb, depth));
      return appendTables([t, ...others]);
    },
    keepTopRows(t, s) {
      t.rows = t.rows.slice(0, Math.max(0, Number(s.count) || 0));
      return t;
    },
    fillDown(t, s) {
      needColumn(t, s.column);
      let last = null;
      t.rows.forEach((r) => { if (blank(r[s.column])) r[s.column] = last; else last = r[s.column]; });
      return t;
    },
  };

  function uniqueAgainst(names, existing) {
    const taken = new Set(existing);
    return names.map((n) => {
      let x = n, i = 2;
      while (taken.has(x)) x = n + '.' + i++;
      taken.add(x);
      return x;
    });
  }

  // Merge keys: text matches case-sensitively in Power Query, numbers by value.
  function matchKey(v) {
    if (typeof v === 'number') return 'n:' + v;
    if (typeof v === 'string') return 's:' + v;
    return 'x:' + String(v);
  }

  function loadOther(w, wb, depth) {
    const src = typeof w === 'string' ? { kind: 'query', name: w } : w;
    const r = loadSource(src, wb, (depth || 0) + 1);
    if (r.error) throw new Error(r.error);
    // apply the referenced query's own source types when it is a raw file:
    return r.table;
  }

  // Run a query. `upto` (exclusive index) previews the table after N steps.
  function run(def, wb, upto, depth = 0) {
    const src = loadSource(def.source, wb, depth);
    if (src.error) return { error: src.error, errorStep: -1 };
    let t = cloneTable(src.table);
    const steps = def.steps || [];
    const end = upto === undefined ? steps.length : Math.min(upto, steps.length);
    for (let i = 0; i < end; i++) {
      const s = steps[i];
      const fn = STEPS[s.type];
      if (!fn) return { error: 'Unknown step "' + s.type + '".', errorStep: i, table: t };
      try {
        t = fn(cloneTable(t), s, wb, depth) || t;
      } catch (e) {
        return { error: e.message, errorStep: i, table: t };
      }
    }
    const errors = t.rows.reduce((n, r) => n + t.columns.filter((c) => isErr(r[c])).length, 0);
    return { table: t, errorCells: errors };
  }

  const api = { run, loadSource, tableFromText, convert, STEPS, STEP_LABELS, describe, sourceLabel, appendTables, compileCustom };
  root.DIM = root.DIM || {};
  root.DIM.Query = api;
  if (isNode) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
