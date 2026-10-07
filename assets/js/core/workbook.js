/*
 * workbook.js — the spreadsheet model behind every interactive workbook.
 *
 * A Workbook holds sheets of cells. Each cell stores what the student typed
 * (`input`), plus optional `format`, `style` and `locked` flags. Values are
 * computed on demand by the formula engine and cached until the next edit.
 *
 * Also implements the "ribbon" operations the course teaches: fill down/right,
 * copy/paste with relative-reference shifting, sort, text to columns, remove
 * duplicates, CSV import, Excel Tables with structured references, defined
 * names, data validation, conditional formatting, charts, pivot tables and
 * Power Query definitions (the last two delegate to pivot.js / query.js).
 */
(function (root) {
  'use strict';

  const isNode = typeof module === 'object' && module.exports;
  const F = isNode ? require('./formula.js') : root.DIM.Formula;
  const getPivot = () => (isNode ? require('./pivot.js') : root.DIM.Pivot);
  const getQuery = () => (isNode ? require('./query.js') : root.DIM.Query);

  const { parseAddress, parseRangeAddress, toAddress, indexToCol, isError, isBlank, isRange, RangeRef, ERR } = F;

  const MAX_UNDO = 100;

  // ---------------------------------------------------------------------------
  // Input parsing: what does a typed string become?
  // ---------------------------------------------------------------------------

  // Returns { kind: 'formula'|'value', value, autoFormat }.
  function interpretInput(input, format) {
    if (input === null || input === undefined || input === '') return { kind: 'value', value: null };
    if (typeof input === 'number') return { kind: 'value', value: input };
    if (typeof input === 'boolean') return { kind: 'value', value: input };
    const s = String(input);
    if (format === 'text') return { kind: 'value', value: s.startsWith("'") ? s.slice(1) : s };
    if (s.startsWith('=') && s.length > 1) return { kind: 'formula' };
    if (s.startsWith("'")) return { kind: 'value', value: s.slice(1) };
    const up = s.trim().toUpperCase();
    if (up === 'TRUE' || up === 'FALSE') return { kind: 'value', value: up === 'TRUE' };
    const err = F.ERR && Object.values(F.ERR).find((c) => c === up);
    if (err) return { kind: 'value', value: F.err(err) };
    const p = F.parseNumberLike(s);
    if (p) return { kind: 'value', value: p.value, autoFormat: p.format };
    return { kind: 'value', value: s };
  }

  // The string to put in a cell so that it round-trips to `value` exactly.
  function inputForValue(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number' || typeof value === 'boolean') return value;
    if (isError(value)) return value.code;
    const s = String(value);
    const it = interpretInput(s);
    if (it.kind === 'formula' || typeof it.value !== 'string' || it.value !== s) return "'" + s;
    return s;
  }

  function cellKey(row, col) { return toAddress(row, col); }

  // "[x]" -> "x" only when the first bracket closes at the very end
  // ("[#This Row],[Qty]" is left alone).
  function stripBrackets(text) {
    const t = String(text).trim();
    if (!t.startsWith('[') || !t.endsWith(']')) return t;
    let depth = 0;
    for (let i = 0; i < t.length; i++) {
      if (t[i] === '[') depth++;
      else if (t[i] === ']') { depth--; if (depth === 0 && i < t.length - 1) return t; }
    }
    return t.slice(1, -1).trim();
  }

  function* eachCell(range) {
    for (let r = range.r1; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) yield [r, c];
  }

  function normRange(range) {
    if (typeof range === 'string') {
      const r = parseRangeAddress(range);
      if (!r) throw new Error('Bad range: ' + range);
      return r;
    }
    return range;
  }

  function deepClone(x) { return JSON.parse(JSON.stringify(x)); }

  // ---------------------------------------------------------------------------
  // Workbook
  // ---------------------------------------------------------------------------

  class Workbook {
    constructor() {
      this.sheets = [];
      this.tables = [];       // { name, sheet, ref }
      this.names = {};        // NAME -> 'Sheet!$A$1:$A$9'
      this.charts = [];       // { id, sheet, type, range, title, axisMin }
      this.pivots = [];       // { id, sourceSheet, sourceRange, rows, columns, values, agg, targetSheet }
      this.queries = [];      // Power Query definitions (see query.js)
      this.sources = {};      // "files" available to Get Data: name -> text
      this._cache = new Map();
      this._evaluating = new Set();
      this._undo = [];
      this._redo = [];
      this.version = 0;
      this.listeners = new Set();
    }

    // ----- construction ------------------------------------------------------

    static fromSpec(spec) {
      const wb = new Workbook();
      (spec.sheets || []).forEach((s) => wb._loadSheetSpec(s));
      if (!wb.sheets.length) wb.addSheet('Sheet1', { silent: true });
      (spec.tables || []).forEach((t) => wb.tables.push({ name: t.name, sheet: t.sheet || wb.sheets[0].name, ref: t.ref }));
      Object.assign(wb.names, spec.names || {});
      (spec.charts || []).forEach((c) => wb.charts.push({ id: c.id || wb._id('chart'), ...c }));
      Object.assign(wb.sources, spec.sources || {});
      (spec.queries || []).forEach((q) => wb.queries.push(deepClone(q)));
      (spec.pivots || []).forEach((p) => wb.pivots.push(deepClone(p)));
      wb.recalc();
      return wb;
    }

    _loadSheetSpec(s) {
      const sheet = this.addSheet(s.name, { silent: true });
      const rows = s.rows || [];
      rows.forEach((row, r) => {
        (row || []).forEach((v, c) => {
          if (v === null || v === undefined || v === '') return;
          sheet.cells[cellKey(r, c)] = { input: v };
        });
      });
      for (const [addr, spec] of Object.entries(s.cells || {})) {
        const a = parseAddress(addr);
        const key = cellKey(a.row, a.col);
        sheet.cells[key] = typeof spec === 'object' && spec !== null && !Array.isArray(spec)
          ? { ...(sheet.cells[key] || {}), ...spec }
          : { ...(sheet.cells[key] || {}), input: spec };
      }
      for (const [rng, fmt] of Object.entries(s.formats || {})) {
        for (const [r, c] of eachCell(normRange(rng))) {
          const key = cellKey(r, c);
          sheet.cells[key] = { ...(sheet.cells[key] || { input: '' }), format: fmt };
        }
      }
      for (const [rng, style] of Object.entries(s.styles || {})) {
        for (const [r, c] of eachCell(normRange(rng))) {
          const key = cellKey(r, c);
          sheet.cells[key] = { ...(sheet.cells[key] || { input: '' }), style: { ...((sheet.cells[key] || {}).style || {}), ...style } };
        }
      }
      if (s.headerRow !== false && rows.length) {
        // first row is a header row: bold by default
        (rows[0] || []).forEach((v, c) => {
          const key = cellKey(0, c);
          if (sheet.cells[key] && s.headerRow !== false && typeof v === 'string' && !String(v).startsWith('=')) {
            sheet.cells[key].style = { bold: true, ...(sheet.cells[key].style || {}) };
          }
        });
      }
      if (s.lock) {
        const ranges = s.lock === true ? null : [].concat(s.lock).map(normRange);
        for (const [key, cell] of Object.entries(sheet.cells)) {
          if (cell.input === '' || cell.input === undefined) continue;
          if (!ranges) { cell.locked = true; continue; }
          const a = parseAddress(key);
          if (ranges.some((r) => a.row >= r.r1 && a.row <= r.r2 && a.col >= r.c1 && a.col <= r.c2)) cell.locked = true;
        }
      }
      sheet.widths = { ...(s.widths || {}) };
      sheet.validations = deepClone(s.validations || []);
      sheet.condFormats = deepClone(s.condFormats || []);
      if (s.notes) sheet.notes = { ...s.notes };
      return sheet;
    }

    toJSON() {
      return {
        version: 1,
        sheets: this.sheets.map((s) => ({
          name: s.name, cells: s.cells, widths: s.widths, validations: s.validations, condFormats: s.condFormats, notes: s.notes,
        })),
        tables: this.tables, names: this.names, charts: this.charts, pivots: this.pivots,
        queries: this.queries, sources: this.sources,
      };
    }

    static fromJSON(json) {
      const wb = new Workbook();
      const data = typeof json === 'string' ? JSON.parse(json) : deepClone(json);
      wb.sheets = (data.sheets || []).map((s) => ({
        name: s.name, cells: s.cells || {}, widths: s.widths || {}, validations: s.validations || [],
        condFormats: s.condFormats || [], notes: s.notes,
      }));
      wb.tables = data.tables || [];
      wb.names = data.names || {};
      wb.charts = data.charts || [];
      wb.pivots = data.pivots || [];
      wb.queries = data.queries || [];
      wb.sources = data.sources || {};
      if (!wb.sheets.length) wb.addSheet('Sheet1', { silent: true });
      wb.recalc();
      return wb;
    }

    clone() { return Workbook.fromJSON(this.toJSON()); }

    // ----- change tracking & undo ---------------------------------------------

    onChange(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }

    _snapshot() { return JSON.stringify(this.toJSON()); }

    _restore(snap) {
      const other = Workbook.fromJSON(snap);
      for (const k of ['sheets', 'tables', 'names', 'charts', 'pivots', 'queries', 'sources']) this[k] = other[k];
    }

    // Wrap a mutation so it can be undone and listeners hear about it.
    transact(label, fn) {
      const before = this._snapshot();
      const result = fn();
      if (result === false) return false;
      this._undo.push({ label, snap: before });
      if (this._undo.length > MAX_UNDO) this._undo.shift();
      this._redo = [];
      this.recalc();
      this._emit(label);
      return result === undefined ? true : result;
    }

    undo() {
      const step = this._undo.pop();
      if (!step) return false;
      this._redo.push({ label: step.label, snap: this._snapshot() });
      this._restore(step.snap);
      this.recalc();
      this._emit('undo');
      return true;
    }

    redo() {
      const step = this._redo.pop();
      if (!step) return false;
      this._undo.push({ label: step.label, snap: this._snapshot() });
      this._restore(step.snap);
      this.recalc();
      this._emit('redo');
      return true;
    }

    get canUndo() { return this._undo.length > 0; }
    get canRedo() { return this._redo.length > 0; }

    _emit(label) {
      this.version++;
      for (const fn of this.listeners) { try { fn(label); } catch (e) { /* listener errors never break the model */ } }
    }

    _id(prefix) { return prefix + '-' + Math.random().toString(36).slice(2, 8); }

    // ----- sheets --------------------------------------------------------------

    getSheet(name) {
      if (name === undefined || name === null) return this.sheets[0];
      const lower = String(name).toLowerCase();
      return this.sheets.find((s) => s.name.toLowerCase() === lower) || null;
    }

    addSheet(name, opts = {}) {
      let base = name || 'Sheet' + (this.sheets.length + 1);
      let unique = base, i = 2;
      while (this.getSheet(unique)) unique = base + ' (' + i++ + ')';
      const sheet = { name: unique, cells: {}, widths: {}, validations: [], condFormats: [] };
      if (opts.silent) { this.sheets.push(sheet); return sheet; }
      this.transact('Add sheet', () => { this.sheets.push(sheet); });
      return this.getSheet(unique);
    }

    renameSheet(oldName, newName) {
      if (!newName || this.getSheet(newName)) return false;
      return this.transact('Rename sheet', () => {
        const s = this.getSheet(oldName);
        if (!s) return false;
        s.name = newName;
        this.tables.forEach((t) => { if (t.sheet === oldName) t.sheet = newName; });
        this.charts.forEach((c) => { if (c.sheet === oldName) c.sheet = newName; });
      });
    }

    removeSheet(name) {
      if (this.sheets.length <= 1) return false;
      return this.transact('Delete sheet', () => {
        const i = this.sheets.findIndex((s) => s.name === name);
        if (i < 0) return false;
        this.sheets.splice(i, 1);
        this.tables = this.tables.filter((t) => t.sheet !== name);
        this.charts = this.charts.filter((c) => c.sheet !== name);
      });
    }

    usedBounds(sheetName) {
      const s = this.getSheet(sheetName);
      let rows = 0, cols = 0;
      if (s) {
        for (const [key, cell] of Object.entries(s.cells)) {
          if (cell.input === '' || cell.input === undefined || cell.input === null) continue;
          const a = parseAddress(key);
          rows = Math.max(rows, a.row + 1);
          cols = Math.max(cols, a.col + 1);
        }
      }
      return { rows, cols };
    }

    // ----- cell access -----------------------------------------------------------

    _cell(sheetName, row, col) {
      const s = this.getSheet(sheetName);
      return s ? s.cells[cellKey(row, col)] : undefined;
    }

    getCell(sheetName, addr) {
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      return this._cell(sheetName, a.row, a.col) || null;
    }

    getInput(sheetName, addr) {
      const c = this.getCell(sheetName, addr);
      return c && c.input !== undefined && c.input !== null ? String(c.input) : '';
    }

    getFormat(sheetName, addr) {
      const c = this.getCell(sheetName, addr);
      if (!c) return null;
      if (c.format) return c.format;
      const it = interpretInput(c.input);
      return it.autoFormat || null;
    }

    isFormula(sheetName, addr) {
      return this.getInput(sheetName, addr).startsWith('=') && this.getInput(sheetName, addr).length > 1;
    }

    recalc() { this._cache.clear(); }

    // Evaluated value of a cell (null for blank).
    getValue(sheetName, addr) {
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      const s = this.getSheet(sheetName);
      if (!s || !a) return F.err(ERR.REF);
      return this._valueAt(s.name, a.row, a.col);
    }

    _valueAt(sheetName, row, col) {
      const s = this.getSheet(sheetName);
      if (!s) return F.err(ERR.REF, 'Sheet "' + sheetName + '" does not exist');
      const key = s.name + '!' + row + ',' + col;
      if (this._cache.has(key)) return this._cache.get(key);
      const cell = s.cells[cellKey(row, col)];
      if (!cell || cell.input === '' || cell.input === undefined || cell.input === null) return null;
      const it = interpretInput(cell.input, cell.format);
      if (it.kind === 'value') {
        this._cache.set(key, it.value);
        return it.value;
      }
      if (this._evaluating.has(key)) return F.err(ERR.CIRC, 'Circular reference: this formula refers to itself');
      this._evaluating.add(key);
      let value;
      try {
        const ast = this._parsed(cell);
        if (ast instanceof Error) value = F.err(ERR.NAME, ast.message);
        else {
          const ctx = this._ctx(s.name, row, col);
          value = F.deref(F.evaluate(ast, ctx), ctx);
          if (isRange(value)) value = F.err(ERR.VALUE);
          if (value === null || value === undefined) value = 0; // =A1 of a blank shows 0, as in Excel
          if (typeof value === 'number' && !Number.isFinite(value)) value = F.err(ERR.NUM);
        }
      } catch (e) {
        value = F.err(ERR.VALUE, e.message);
      } finally {
        this._evaluating.delete(key);
      }
      // Uploaded files: if our engine lacks a function, trust Excel's cached result.
      if (isError(value) && value.code === ERR.NAME && cell.cached !== undefined) value = cell.cached;
      this._cache.set(key, value);
      return value;
    }

    _parsed(cell) {
      if (cell._astFor === cell.input && cell._ast) return cell._ast;
      let ast;
      try { ast = F.parse(cell.input); } catch (e) { ast = e; }
      Object.defineProperty(cell, '_ast', { value: ast, writable: true, configurable: true, enumerable: false });
      Object.defineProperty(cell, '_astFor', { value: cell.input, writable: true, configurable: true, enumerable: false });
      return ast;
    }

    _ctx(sheetName, row, col) {
      return {
        sheet: sheetName, row, col,
        getCell: (sh, r, c) => this._valueAt(sh, r, c),
        usedBounds: (sh) => this.usedBounds(sh),
        resolveName: (name) => this.resolveName(name),
        resolveTable: (table, spec, ctx) => this.resolveTable(table, spec, ctx),
      };
    }

    // Parse error message for a formula cell, if any.
    formulaError(sheetName, addr) {
      const cell = this.getCell(sheetName, addr);
      if (!cell || !String(cell.input || '').startsWith('=')) return null;
      const ast = this._parsed(cell);
      return ast instanceof Error ? ast.message : null;
    }

    display(sheetName, addr) {
      const v = this.getValue(sheetName, addr);
      return F.formatValue(v, this.getFormat(sheetName, addr));
    }

    // Values of a range as a 2D array.
    getRangeValues(sheetName, range) {
      const r = normRange(range);
      const out = [];
      for (let i = r.r1; i <= r.r2; i++) {
        const row = [];
        for (let j = r.c1; j <= r.c2; j++) row.push(this._valueAt(sheetName, i, j));
        out.push(row);
      }
      return out;
    }

    // ----- editing ---------------------------------------------------------------

    isLocked(sheetName, addr) {
      const c = this.getCell(sheetName, addr);
      return !!(c && c.locked);
    }

    // Validate a value against data-validation rules. Returns null or a message.
    validate(sheetName, addr, input) {
      const s = this.getSheet(sheetName);
      if (!s || !s.validations || !s.validations.length || input === '' || String(input).startsWith('=')) return null;
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      for (const v of s.validations) {
        const r = normRange(v.range);
        if (a.row < r.r1 || a.row > r.r2 || a.col < r.c1 || a.col > r.c2) continue;
        const it = interpretInput(input);
        if (v.type === 'list') {
          const ok = (v.values || []).some((x) => String(x).toLowerCase() === String(it.value).toLowerCase());
          if (!ok) return v.message || 'Choose one of: ' + v.values.join(', ');
        } else if (v.type === 'whole' || v.type === 'decimal') {
          const n = it.value;
          if (typeof n !== 'number' || (v.type === 'whole' && !Number.isInteger(n))) return v.message || 'Enter a ' + (v.type === 'whole' ? 'whole number' : 'number');
          if ((v.min !== undefined && n < v.min) || (v.max !== undefined && n > v.max)) {
            return v.message || 'Enter a number between ' + v.min + ' and ' + v.max;
          }
        } else if (v.type === 'textLength') {
          const len = String(it.value ?? '').length;
          if ((v.min !== undefined && len < v.min) || (v.max !== undefined && len > v.max)) return v.message || 'Text length must be ' + v.min + '–' + v.max;
        }
      }
      return null;
    }

    setInput(sheetName, addr, input, opts = {}) {
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      if (!opts.force && this.isLocked(sheetName, a)) return false;
      return this.transact('Edit ' + toAddress(a.row, a.col), () => this._setRaw(sheetName, a.row, a.col, input));
    }

    _setRaw(sheetName, row, col, input) {
      const s = this.getSheet(sheetName);
      if (!s) return false;
      const key = cellKey(row, col);
      const prev = s.cells[key];
      if (input === '' || input === null || input === undefined) {
        if (prev) {
          if (prev.format || prev.style || prev.locked) { prev.input = ''; delete prev.cached; }
          else delete s.cells[key];
        }
      } else {
        s.cells[key] = { ...(prev || {}), input };
        delete s.cells[key].cached;
        this._maybeExpandTable(s.name, row, col);
      }
      return true;
    }

    _maybeExpandTable(sheetName, row, col) {
      for (const t of this.tables) {
        if (t.sheet !== sheetName) continue;
        const r = normRange(t.ref);
        if (row === r.r2 + 1 && col >= r.c1 && col <= r.c2) t.ref = F.rangeToAddress({ ...r, r2: row });
        // typing a header just right of a table adds a column, as in Excel
        else if (row === r.r1 && col === r.c2 + 1 && !this.tableAt(sheetName, row, col)) t.ref = F.rangeToAddress({ ...r, c2: col });
      }
    }

    clearRange(sheetName, range) {
      const r = normRange(range);
      return this.transact('Clear', () => {
        let changed = false;
        for (const [i, j] of eachCell(r)) {
          if (this._cell(sheetName, i, j) && !this._cell(sheetName, i, j).locked) { this._setRaw(sheetName, i, j, ''); changed = true; }
        }
        return changed ? true : false;
      });
    }

    setFormat(sheetName, range, format) {
      const r = normRange(range);
      const s = this.getSheet(sheetName);
      return this.transact('Format', () => {
        for (const [i, j] of eachCell(r)) {
          const key = cellKey(i, j);
          const cell = s.cells[key] || (s.cells[key] = { input: '' });
          if (format === 'general' || !format) delete cell.format; else cell.format = format;
        }
      });
    }

    setStyle(sheetName, range, style) {
      const r = normRange(range);
      const s = this.getSheet(sheetName);
      return this.transact('Style', () => {
        for (const [i, j] of eachCell(r)) {
          const key = cellKey(i, j);
          const cell = s.cells[key] || (s.cells[key] = { input: '' });
          cell.style = { ...(cell.style || {}), ...style };
          for (const k of Object.keys(cell.style)) if (cell.style[k] === null || cell.style[k] === false) delete cell.style[k];
        }
      });
    }

    // Fill the first row/column of `range` across the rest of it (Ctrl+D / Ctrl+R).
    fill(sheetName, range, direction = 'down') {
      const r = normRange(range);
      return this.transact(direction === 'down' ? 'Fill down' : 'Fill right', () => {
        let changed = false;
        if (direction === 'down') {
          for (let c = r.c1; c <= r.c2; c++) {
            const src = this._cell(sheetName, r.r1, c);
            for (let row = r.r1 + 1; row <= r.r2; row++) {
              const tgt = this._cell(sheetName, row, c);
              if (tgt && tgt.locked) continue;
              this._copyCell(sheetName, r.r1, c, sheetName, row, c, src);
              changed = true;
            }
          }
        } else {
          for (let row = r.r1; row <= r.r2; row++) {
            const src = this._cell(sheetName, row, r.c1);
            for (let c = r.c1 + 1; c <= r.c2; c++) {
              const tgt = this._cell(sheetName, row, c);
              if (tgt && tgt.locked) continue;
              this._copyCell(sheetName, row, r.c1, sheetName, row, c, src);
              changed = true;
            }
          }
        }
        return changed;
      });
    }

    _copyCell(fromSheet, fr, fc, toSheet, tr, tc, srcCell) {
      const src = srcCell === undefined ? this._cell(fromSheet, fr, fc) : srcCell;
      const s = this.getSheet(toSheet);
      const key = cellKey(tr, tc);
      if (!src) { if (s.cells[key] && !s.cells[key].locked) delete s.cells[key]; return; }
      const input = typeof src.input === 'string' && src.input.startsWith('=')
        ? F.shiftFormula(src.input, tr - fr, tc - fc)
        : src.input;
      const next = { input };
      if (src.format) next.format = src.format;
      if (src.style) next.style = { ...src.style };
      s.cells[key] = next;
    }

    // Internal clipboard: capture a block of cells.
    copy(sheetName, range) {
      const r = normRange(range);
      const cells = [];
      for (let i = r.r1; i <= r.r2; i++) {
        const row = [];
        for (let j = r.c1; j <= r.c2; j++) {
          const c = this._cell(sheetName, i, j);
          row.push(c ? { input: c.input, format: c.format, style: c.style } : null);
        }
        cells.push(row);
      }
      return { sheet: sheetName, r1: r.r1, c1: r.c1, cells };
    }

    paste(sheetName, addr, clip, opts = {}) {
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      return this.transact('Paste', () => {
        clip.cells.forEach((row, i) => row.forEach((cell, j) => {
          const tr = a.row + i, tc = a.col + j;
          const tgt = this._cell(sheetName, tr, tc);
          if (tgt && tgt.locked) return;
          if (!cell) { this._setRaw(sheetName, tr, tc, ''); return; }
          let input = cell.input;
          if (opts.valuesOnly) input = inputForValue(this._valueAt(clip.sheet, clip.r1 + i, clip.c1 + j));
          else if (typeof input === 'string' && input.startsWith('=')) input = F.shiftFormula(input, tr - (clip.r1 + i), tc - (clip.c1 + j));
          const s = this.getSheet(sheetName);
          s.cells[cellKey(tr, tc)] = { input, ...(cell.format ? { format: cell.format } : {}), ...(cell.style ? { style: { ...cell.style } } : {}) };
        }));
      });
    }

    // Paste tab-separated text (what Excel/Sheets put on the clipboard).
    pasteText(sheetName, addr, text) {
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      const rows = String(text).replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
      return this.transact('Paste', () => {
        rows.forEach((row, i) => row.forEach((v, j) => {
          const tgt = this._cell(sheetName, a.row + i, a.col + j);
          if (tgt && tgt.locked) return;
          this._setRaw(sheetName, a.row + i, a.col + j, v);
        }));
      });
    }

    // Find the contiguous block around a cell (Excel's "current region").
    currentRegion(sheetName, addr) {
      const a = typeof addr === 'string' ? parseAddress(addr) : addr;
      const filled = (r, c) => r >= 0 && c >= 0 && this.getInput(sheetName, { row: r, col: c }) !== '';
      let r1 = a.row, r2 = a.row, c1 = a.col, c2 = a.col;
      let grew = true;
      while (grew) {
        grew = false;
        const rowHas = (r) => { for (let c = c1 - 1; c <= c2 + 1; c++) if (filled(r, c)) return true; return false; };
        const colHas = (c) => { for (let r = r1 - 1; r <= r2 + 1; r++) if (filled(r, c)) return true; return false; };
        if (r1 > 0 && rowHas(r1 - 1)) { r1--; grew = true; }
        if (rowHas(r2 + 1)) { r2++; grew = true; }
        if (c1 > 0 && colHas(c1 - 1)) { c1--; grew = true; }
        if (colHas(c2 + 1)) { c2++; grew = true; }
        if (r2 - r1 > 5000 || c2 - c1 > 500) break;
      }
      return { r1, c1, r2, c2 };
    }

    // Sort rows of `range` by one of its columns (0-based offset within range).
    sortRange(sheetName, range, colOffset, ascending = true, hasHeader = true) {
      const r = normRange(range);
      const s = this.getSheet(sheetName);
      const start = hasHeader ? r.r1 + 1 : r.r1;
      for (let i = start; i <= r.r2; i++) for (let j = r.c1; j <= r.c2; j++) {
        const c = s.cells[cellKey(i, j)];
        if (c && c.locked && !c.sortable) return { error: 'Some cells in this range are locked.' };
      }
      return this.transact('Sort', () => {
        const rows = [];
        for (let i = start; i <= r.r2; i++) {
          const cells = [];
          for (let j = r.c1; j <= r.c2; j++) cells.push(s.cells[cellKey(i, j)] ? deepClone(s.cells[cellKey(i, j)]) : null);
          rows.push({ key: this._valueAt(sheetName, i, r.c1 + colOffset), cells, from: i });
        }
        rows.sort((x, y) => {
          const xb = isBlank(x.key) || x.key === '', yb = isBlank(y.key) || y.key === '';
          if (xb !== yb) return xb ? 1 : -1; // blanks always last
          const c = F.compare(x.key, y.key);
          return (ascending ? c : -c) || x.from - y.from;
        });
        rows.forEach((row, idx) => {
          const target = start + idx;
          row.cells.forEach((cell, j) => {
            const key = cellKey(target, r.c1 + j);
            if (!cell) { delete s.cells[key]; return; }
            if (typeof cell.input === 'string' && cell.input.startsWith('=')) cell.input = F.shiftFormula(cell.input, target - row.from, 0);
            s.cells[key] = cell;
          });
        });
      });
    }

    // Split one column on a delimiter into the columns to its right.
    textToColumns(sheetName, range, delimiter, opts = {}) {
      const r = normRange(range);
      if (r.c1 !== r.c2) return { error: 'Select a single column to split.' };
      const delim = delimiter === 'tab' ? '\t' : delimiter === 'space' ? ' ' : delimiter;
      const textCols = new Set(opts.textColumns || []);
      return this.transact('Text to Columns', () => {
        for (let i = r.r1; i <= r.r2; i++) {
          const v = this._valueAt(sheetName, i, r.c1);
          if (isBlank(v)) continue;
          let parts = splitDelimited(String(F.toText(v)), delim);
          if (opts.trim !== false) parts = parts.map((p) => p.trim());
          parts.forEach((p, j) => {
            const input = textCols.has(j) ? (p === '' ? '' : "'" + p) : p;
            const s = this.getSheet(sheetName);
            const key = cellKey(i, r.c1 + j);
            const keep = s.cells[key] && s.cells[key].style ? { style: s.cells[key].style } : {};
            s.cells[key] = { ...keep, input };
          });
        }
      });
    }

    removeDuplicates(sheetName, range, hasHeader = true, columns = null) {
      const r = normRange(range);
      const s = this.getSheet(sheetName);
      const start = hasHeader ? r.r1 + 1 : r.r1;
      let removed = 0;
      const ok = this.transact('Remove duplicates', () => {
        const seen = new Set();
        const keep = [];
        for (let i = start; i <= r.r2; i++) {
          const vals = [];
          for (let j = r.c1; j <= r.c2; j++) {
            if (columns && !columns.includes(j - r.c1)) continue;
            const v = this._valueAt(sheetName, i, j);
            vals.push(typeof v === 'string' ? v.toLowerCase() : String(v));
          }
          const k = vals.join('\u0001');
          if (seen.has(k)) { removed++; continue; }
          seen.add(k);
          const cells = [];
          for (let j = r.c1; j <= r.c2; j++) cells.push(s.cells[cellKey(i, j)] ? deepClone(s.cells[cellKey(i, j)]) : null);
          keep.push({ cells, from: i });
        }
        for (let i = start; i <= r.r2; i++) for (let j = r.c1; j <= r.c2; j++) delete s.cells[cellKey(i, j)];
        keep.forEach(({ cells, from }, idx) => cells.forEach((cell, j) => {
          if (!cell) return;
          // rows move up like deleted rows in Excel: relative references move with them
          if (typeof cell.input === 'string' && cell.input.startsWith('=')) cell.input = F.shiftFormula(cell.input, start + idx - from, 0);
          s.cells[cellKey(start + idx, r.c1 + j)] = cell;
        }));
      });
      return ok ? { removed, remaining: r.r2 - start + 1 - removed } : ok;
    }

    // Import delimited text into a new sheet.
    importDelimited(text, opts = {}) {
      const delim = opts.delimiter === 'tab' ? '\t' : opts.delimiter || detectDelimiter(text);
      const rows = parseDelimited(text, delim);
      const textCols = new Set(opts.textColumns || []);
      const name = opts.sheetName || 'Import';
      let sheet;
      this.transact('Import', () => {
        const base = name; let unique = base, n = 2;
        while (this.getSheet(unique)) unique = base + ' (' + n++ + ')';
        sheet = { name: unique, cells: {}, widths: {}, validations: [], condFormats: [] };
        rows.forEach((row, i) => row.forEach((v, j) => {
          if (v === '') return;
          const asText = textCols.has(j) && i > 0;
          sheet.cells[cellKey(i, j)] = { input: asText ? "'" + v : v };
          if (i === 0 && opts.header !== false) sheet.cells[cellKey(i, j)].style = { bold: true };
        }));
        this.sheets.push(sheet);
      });
      return sheet;
    }

    // ----- Excel Tables, names --------------------------------------------------

    addTable(sheetName, range, name) {
      const r = normRange(range);
      if (r.r2 <= r.r1) return { error: 'A table needs a header row and at least one data row.' };
      for (const t of this.tables) {
        if (t.sheet !== sheetName) continue;
        const o = normRange(t.ref);
        if (!(r.r2 < o.r1 || r.r1 > o.r2 || r.c2 < o.c1 || r.c1 > o.c2)) return { error: 'This range overlaps table ' + t.name + '.' };
      }
      let tname = name || 'Table' + (this.tables.length + 1);
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(tname)) return { error: 'Table names must start with a letter and contain no spaces.' };
      if (this.tables.some((t) => t.name.toLowerCase() === tname.toLowerCase())) return { error: 'There is already a table named ' + tname + '.' };
      this.transact('Format as Table', () => {
        this.tables.push({ name: tname, sheet: sheetName, ref: F.rangeToAddress(r) });
        for (let c = r.c1; c <= r.c2; c++) {
          const key = cellKey(r.r1, c);
          const s = this.getSheet(sheetName);
          const cell = s.cells[key] || (s.cells[key] = { input: 'Column' + (c - r.c1 + 1) });
          cell.style = { ...(cell.style || {}), bold: true };
        }
      });
      return { name: tname };
    }

    tableAt(sheetName, row, col) {
      return this.tables.find((t) => {
        if (t.sheet !== sheetName) return false;
        const r = normRange(t.ref);
        return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
      }) || null;
    }

    tableColumns(t) {
      const r = normRange(t.ref);
      const cols = [];
      for (let c = r.c1; c <= r.c2; c++) cols.push(String(F.toText(this._valueAt(t.sheet, r.r1, c))));
      return cols;
    }

    resolveTable(tableName, spec, ctx) {
      const t = tableName
        ? this.tables.find((x) => x.name.toLowerCase() === tableName.toLowerCase())
        : this.tableAt(ctx.sheet, ctx.row, ctx.col);
      if (!t) return null;
      const r = normRange(t.ref);
      // forms typed in the app: Column | @Column | [@Column] | [#All] | #Data
      // forms Excel saves in files: [#This Row],[Column] | [#All],[Column] | [Col1]:[Col2]
      let s = stripBrackets(spec);
      let thisRow = false, area = 'data';
      if (s.startsWith('@')) { thisRow = true; s = s.slice(1).trim(); }
      let colSpec = null;
      for (const piece of s.split(/\s*,\s*(?=\[|#)/)) {
        const p = piece.trim();
        const item = p.replace(/^\[([^\]]*)\]$/, '$1').trim();
        if (/^#This Row$/i.test(item)) thisRow = true;
        else if (/^#All$/i.test(item)) area = 'all';
        else if (/^#Data$/i.test(item)) area = 'data';
        else if (/^#Headers$/i.test(item)) area = 'headers';
        else if (/^#Totals$/i.test(item)) return null;
        else if (p !== '') colSpec = p;
      }
      let c1 = r.c1, c2 = r.c2;
      if (colSpec !== null) {
        const cols = this.tableColumns(t).map((x) => x.toLowerCase());
        const ends = colSpec.split(/\]\s*:\s*\[/).map((x) => x.replace(/^\[|\]$/g, '').trim().toLowerCase());
        const i1 = cols.indexOf(ends[0]), i2 = cols.indexOf(ends[ends.length - 1]);
        if (i1 < 0 || i2 < 0) return null;
        c1 = r.c1 + Math.min(i1, i2); c2 = r.c1 + Math.max(i1, i2);
      }
      if (thisRow) {
        if (ctx.row <= r.r1 || ctx.row > r.r2) return F.err(ERR.VALUE);
        return new RangeRef(t.sheet, ctx.row, c1, ctx.row, c2);
      }
      if (area === 'headers') return new RangeRef(t.sheet, r.r1, c1, r.r1, c2);
      return new RangeRef(t.sheet, area === 'all' ? r.r1 : r.r1 + 1, c1, r.r2, c2);
    }

    defineName(name, ref) {
      if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || parseAddress(name)) return { error: 'Names must start with a letter, have no spaces, and not look like a cell address.' };
      this.transact('Define name', () => { this.names[name] = ref; });
      return { name };
    }

    resolveName(name) {
      const key = Object.keys(this.names).find((k) => k.toLowerCase() === String(name).toLowerCase());
      if (!key) return null;
      const r = parseRangeAddress(this.names[key]);
      if (!r) return null;
      return new RangeRef(r.sheet || this.sheets[0].name, r.r1, r.c1, r.r2, r.c2);
    }

    // Rewrite structured references and names as plain A1 references (for xlsx export).
    // Tables listed in opts.keepTables stay as structured references, in the
    // form Excel saves them (Name[Col], Name[[#This Row],[Col]]).
    formulaToA1(sheetName, row, col, input, opts = {}) {
      if (!String(input).startsWith('=')) return input;
      let tokens;
      try { tokens = F.tokenize(String(input).slice(1)); } catch (e) { return input; }
      const ctx = { sheet: sheetName, row, col };
      const keep = opts.keepTables;
      let out = '=';
      for (const t of tokens) {
        if (t.type === 'table' && keep) {
          const tbl = t.table ? this.tables.find((x) => x.name.toLowerCase() === t.table.toLowerCase()) : this.tableAt(sheetName, row, col);
          if (tbl && keep.has(tbl.name)) {
            const spec = stripBrackets(t.spec);
            if (spec.startsWith('@')) {
              const c = stripBrackets(spec.slice(1));
              out += tbl.name + '[' + (c ? '[#This Row],[' + c + ']' : '#This Row') + ']';
            } else out += tbl.name + '[' + spec + ']';
            continue;
          }
        }
        if (t.type === 'table') {
          const rr = this.resolveTable(t.table, t.spec, ctx);
          if (rr && isRange(rr)) {
            const prefix = rr.sheet !== sheetName ? F.quoteSheet(rr.sheet) + '!' : '';
            const abs = !t.spec.includes('@');
            const a = toAddress(rr.r1, rr.c1, abs, abs), b = toAddress(rr.r2, rr.c2, abs, abs);
            out += prefix + (a === b ? a : a + ':' + b);
            continue;
          }
        }
        out += t.text;
      }
      return out;
    }

    // ----- validation & conditional formatting -----------------------------------

    addValidation(sheetName, range, rule) {
      return this.transact('Data validation', () => {
        const s = this.getSheet(sheetName);
        s.validations = (s.validations || []).filter((v) => v.range !== range);
        s.validations.push({ range, ...rule });
      });
    }

    addCondFormat(sheetName, range, rule) {
      return this.transact('Conditional formatting', () => {
        const s = this.getSheet(sheetName);
        s.condFormats = s.condFormats || [];
        s.condFormats.push({ range, ...rule });
      });
    }

    clearCondFormats(sheetName) {
      return this.transact('Clear rules', () => { this.getSheet(sheetName).condFormats = []; });
    }

    // Which conditional-format style (if any) applies to a cell.
    condStyle(sheetName, row, col) {
      const s = this.getSheet(sheetName);
      if (!s || !s.condFormats || !s.condFormats.length) return null;
      const v = this._valueAt(sheetName, row, col);
      for (const cf of s.condFormats) {
        const r = normRange(cf.range);
        if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
        if (condMatches(cf, v, () => this.getRangeValues(sheetName, r).flat())) return cf.style || 'green';
      }
      return null;
    }

    // ----- charts -----------------------------------------------------------------

    addChart(def) {
      const chart = { id: this._id('chart'), title: '', axisMin: null, ...def };
      this.transact('Insert chart', () => { this.charts.push(chart); });
      return chart;
    }

    updateChart(id, patch) {
      return this.transact('Edit chart', () => {
        const c = this.charts.find((x) => x.id === id);
        if (!c) return false;
        Object.assign(c, patch);
      });
    }

    removeChart(id) {
      return this.transact('Delete chart', () => { this.charts = this.charts.filter((c) => c.id !== id); });
    }

    // Categories + numeric series for a chart, read from its range.
    chartData(chart) {
      const r = normRange(chart.range);
      const sheet = chart.dataSheet || chart.sheet;
      const vals = this.getRangeValues(sheet, r);
      if (!vals.length) return { categories: [], series: [] };
      const firstRowText = vals[0].some((v) => typeof v === 'string');
      const header = firstRowText ? vals[0] : null;
      const body = firstRowText ? vals.slice(1) : vals;
      const firstColText = body.some((row) => typeof row[0] === 'string') || (chart.type === 'scatter' ? false : body.every((row) => typeof row[0] === 'number') && vals[0].length > 1 && header && chart.type !== 'scatter');
      const catIdx = firstColText ? 0 : -1;
      const categories = body.map((row, i) => (catIdx === 0 ? F.formatValue(row[0], this.getFormat(sheet, { row: r.r1 + (header ? 1 : 0) + i, col: r.c1 })) : String(i + 1)));
      const series = [];
      for (let j = catIdx + 1; j < (vals[0] || []).length; j++) {
        series.push({
          name: header ? F.toText(header[j]) : 'Series ' + (j - catIdx),
          values: body.map((row) => (typeof row[j] === 'number' ? row[j] : null)),
          x: chart.type === 'scatter' ? body.map((row) => (typeof row[0] === 'number' ? row[0] : null)) : null,
        });
      }
      if (chart.type === 'scatter' && !firstColText && series.length > 1) {
        // first numeric column is X
        const xs = series.shift();
        series.forEach((s) => { s.x = xs.values; });
      }
      return { categories, series };
    }

    // ----- pivot tables -------------------------------------------------------------

    addPivot(def) {
      const Pivot = getPivot();
      const pivot = { id: this._id('pivot'), agg: 'sum', columns: null, ...def };
      let result;
      this.transact('PivotTable', () => {
        if (!pivot.targetSheet) {
          let n = 1; while (this.getSheet('Pivot' + n)) n++;
          pivot.targetSheet = 'Pivot' + n;
        }
        if (!this.getSheet(pivot.targetSheet)) this.sheets.push({ name: pivot.targetSheet, cells: {}, widths: {}, validations: [], condFormats: [] });
        this.pivots.push(pivot);
        result = this._writePivot(pivot, Pivot);
        if (result && result.error) return false;
      });
      return result && result.error ? result : pivot;
    }

    refreshPivots() {
      const Pivot = getPivot();
      return this.transact('Refresh pivots', () => { this.pivots.forEach((p) => this._writePivot(p, Pivot)); });
    }

    removePivot(id) {
      return this.transact('Delete pivot', () => {
        const p = this.pivots.find((x) => x.id === id);
        if (!p) return false;
        const s = this.getSheet(p.targetSheet);
        if (s) s.cells = {};
        this.pivots = this.pivots.filter((x) => x.id !== id);
      });
    }

    _writePivot(p, Pivot) {
      const records = this.recordsFrom(p.sourceSheet, p.sourceRange);
      if (records.error) return records;
      const out = Pivot.build(records.rows, { rows: p.rows, columns: p.columns, values: p.values, agg: p.agg, rowsGroup: p.rowsGroup, columnsGroup: p.columnsGroup });
      if (out.error) return out;
      const s = this.getSheet(p.targetSheet);
      const anchor = parseAddress(p.targetCell || 'A1');
      // clear previous output
      for (const key of Object.keys(s.cells)) if (s.cells[key].pivot === p.id) delete s.cells[key];
      out.matrix.forEach((row, i) => row.forEach((v, j) => {
        if (v === null || v === undefined || v === '') return;
        const cell = { input: inputForValue(v), pivot: p.id, locked: true };
        if (i === 0 || (out.headerRows && i < out.headerRows) || (out.totalRow && i === out.matrix.length - 1) || j === 0) cell.style = { bold: true };
        if (typeof v === 'number' && p.format) cell.format = p.format;
        s.cells[cellKey(anchor.row + i, anchor.col + j)] = cell;
      }));
      return out;
    }

    // Read a header + rows block as records: { columns, rows: [{col: value}] }.
    recordsFrom(sheetName, range) {
      const s = this.getSheet(sheetName);
      if (!s) return { error: 'Sheet "' + sheetName + '" not found.' };
      const r = range ? normRange(range) : this.currentRegion(sheetName, 'A1');
      const vals = this.getRangeValues(sheetName, r);
      if (vals.length < 2) return { error: 'The source needs a header row and at least one row of data.' };
      const columns = vals[0].map((v, i) => (isBlank(v) || v === '' ? 'Column' + (i + 1) : String(F.toText(v))));
      const rows = vals.slice(1).filter((row) => row.some((v) => !isBlank(v) && v !== '')).map((row) => {
        const o = {};
        columns.forEach((c, i) => { o[c] = isError(row[i]) ? row[i].code : row[i]; });
        return o;
      });
      return { columns, rows };
    }

    // ----- Power Query ------------------------------------------------------------

    saveQuery(def) {
      return this.transact('Save query', () => {
        const i = this.queries.findIndex((q) => q.id === def.id);
        if (i >= 0) this.queries[i] = deepClone(def); else this.queries.push(deepClone(def));
      });
    }

    runQuery(def) {
      // A query read from a real Excel file keeps its M code, which this
      // simulator does not execute; its result is the table Excel loaded.
      if (def.imported) {
        if (!def.loadedSheet) return { error: 'This query is "connection only": load it to a worksheet (Close & Load To… → Table).' };
        const rec = this.recordsFrom(def.loadedSheet, def.loadedRange);
        if (rec.error) return rec;
        const types = {};
        rec.columns.forEach((c) => { types[c] = 'any'; });
        return { table: { columns: rec.columns, rows: rec.rows, types } };
      }
      return getQuery().run(def, this);
    }

    // Close & Load: run the query and write its output to a sheet named after it.
    loadQuery(id) {
      const def = this.queries.find((q) => q.id === id);
      if (!def) return { error: 'Query not found.' };
      if (def.imported) return this.runQuery(def);
      const res = this.runQuery(def);
      if (res.error) return res;
      this.transact('Load query', () => {
        let s = this.getSheet(def.name);
        if (!s) { s = { name: def.name, cells: {}, widths: {}, validations: [], condFormats: [] }; this.sheets.push(s); }
        s.cells = {};
        res.table.columns.forEach((c, j) => { s.cells[cellKey(0, j)] = { input: inputForValue(c), style: { bold: true }, locked: true, query: id }; });
        res.table.rows.forEach((row, i) => res.table.columns.forEach((c, j) => {
          const v = row[c];
          if (v === null || v === undefined || v === '') return;
          const cell = { input: inputForValue(v), locked: true, query: id };
          const type = res.table.types && res.table.types[c];
          if (type === 'date') cell.format = 'date';
          if (type === 'currency') cell.format = 'currency';
          s.cells[cellKey(i + 1, j)] = cell;
        }));
        def.loadedRows = res.table.rows.length;
      });
      return res;
    }

    refreshAll() {
      const results = this.queries.map((q) => this.loadQuery(q.id));
      if (this.pivots.length) this.refreshPivots();
      return results;
    }

    removeQuery(id) {
      return this.transact('Delete query', () => { this.queries = this.queries.filter((q) => q.id !== id); });
    }

    // ----- selection statistics (status bar) ---------------------------------------

    stats(sheetName, range) {
      const vals = this.getRangeValues(sheetName, range).flat();
      const nums = vals.filter((v) => typeof v === 'number');
      const count = vals.filter((v) => !isBlank(v) && v !== '').length;
      const sum = F.round15(nums.reduce((a, b) => a + b, 0));
      return { count, numericCount: nums.length, sum, average: nums.length ? F.round15(sum / nums.length) : null, min: nums.length ? Math.min(...nums) : null, max: nums.length ? Math.max(...nums) : null };
    }
  }

  // ---------------------------------------------------------------------------
  // Conditional formatting rules
  // ---------------------------------------------------------------------------

  function condMatches(cf, v, allValues) {
    if (isBlank(v) || isError(v)) return cf.op === 'blank' ? isBlank(v) : false;
    const target = cf.value;
    switch (cf.op) {
      case '>': return typeof v === 'number' && v > Number(target);
      case '<': return typeof v === 'number' && v < Number(target);
      case '>=': return typeof v === 'number' && v >= Number(target);
      case '<=': return typeof v === 'number' && v <= Number(target);
      case '=': return F.compare(v, typeof v === 'number' ? Number(target) : String(target)) === 0;
      case '<>': return F.compare(v, typeof v === 'number' ? Number(target) : String(target)) !== 0;
      case 'between': return typeof v === 'number' && v >= Number(cf.value) && v <= Number(cf.value2);
      case 'contains': return String(F.toText(v)).toLowerCase().includes(String(target).toLowerCase());
      case 'duplicate': {
        const all = allValues().filter((x) => !isBlank(x));
        return all.filter((x) => F.compare(x, v) === 0 && typeof x === typeof v).length > 1;
      }
      case 'top': {
        const nums = allValues().filter((x) => typeof x === 'number').sort((a, b) => b - a);
        return typeof v === 'number' && v >= nums[Math.min(nums.length, Number(target)) - 1];
      }
      case 'aboveAverage': {
        const nums = allValues().filter((x) => typeof x === 'number');
        const avg = nums.reduce((a, b) => a + b, 0) / (nums.length || 1);
        return typeof v === 'number' && v > avg;
      }
      default: return false;
    }
  }

  // ---------------------------------------------------------------------------
  // Delimited text helpers (CSV/TSV/pipe), RFC-4180 style quoting
  // ---------------------------------------------------------------------------

  function splitDelimited(line, delim) {
    return parseDelimited(line, delim)[0] || [''];
  }

  function parseDelimited(text, delim) {
    const rows = [];
    let row = [], field = '', inQuotes = false, i = 0;
    const s = String(text).replace(/^﻿/, '');
    while (i < s.length) {
      const ch = s[i];
      if (inQuotes) {
        if (ch === '"') {
          if (s[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += ch; i++; continue;
      }
      if (ch === '"' && field === '') { inQuotes = true; i++; continue; }
      if (s.startsWith(delim, i)) { row.push(field); field = ''; i += delim.length; continue; }
      if (ch === '\r') { i++; continue; }
      if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += ch; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function detectDelimiter(text) {
    const lines = String(text).split(/\r?\n/).filter((l) => l.trim() !== '').slice(0, 5);
    const candidates = [',', '\t', '|', ';'];
    let best = ',', bestScore = -1;
    for (const d of candidates) {
      const counts = lines.map((l) => parseDelimited(l, d)[0].length);
      const consistent = counts.every((c) => c === counts[0]);
      const score = consistent && counts[0] > 1 ? counts[0] : 0;
      if (score > bestScore) { best = d; bestScore = score; }
    }
    return best;
  }

  function toDelimited(rows, delim = ',') {
    return rows.map((row) => row.map((v) => {
      const s = v === null || v === undefined ? '' : String(v);
      return /[",\n\r]/.test(s) || s.includes(delim) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(delim)).join('\r\n');
  }

  // Export a sheet as CSV text — shows students exactly what a CSV keeps.
  function sheetToCSV(wb, sheetName, delim = ',') {
    const b = wb.usedBounds(sheetName);
    const rows = [];
    for (let r = 0; r < b.rows; r++) {
      const row = [];
      for (let c = 0; c < b.cols; c++) {
        const v = wb.getValue(sheetName, { row: r, col: c });
        row.push(isBlank(v) ? '' : wb.display(sheetName, { row: r, col: c }));
      }
      rows.push(row);
    }
    return toDelimited(rows, delim);
  }

  const api = {
    Workbook, interpretInput, inputForValue, parseDelimited, detectDelimiter, toDelimited, sheetToCSV,
    splitDelimited, normRange, condMatches, cellKey,
  };

  root.DIM = root.DIM || {};
  root.DIM.Workbook = api;
  if (isNode) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
