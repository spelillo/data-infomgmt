/*
 * sheet-view.js — the interactive, Excel-like workbook used in lessons, quizzes
 * and assignments.
 *
 * Mount with DIM.SheetView.mount(container, workbook, options):
 *   options.tools       which ribbon tools to show (see TOOL_SETS)
 *   options.highlights  [{ sheet, range, label }] task cells to outline
 *   options.fileName    base name for downloads
 *   options.onChange    called after every change (autosave hook)
 *   options.onReset     called when the student clicks Reset (return a fresh Workbook)
 *
 * Keyboard (matches Excel): arrows, Shift+arrows, Ctrl+arrows, Enter/Tab,
 * F2 edit, F4 toggle $ references, Delete, Ctrl+C/X/V, Ctrl+Z/Y, Ctrl+D/R,
 * Ctrl+B, Ctrl+` show formulas, Ctrl+Home.
 */
(function (root) {
  'use strict';

  const F = root.DIM.Formula;
  const WB = root.DIM.Workbook;
  const { h, dialog, toast, field, select, downloadText, readFile } = root.DIM.Dom;

  const FORMAT_OPTIONS = [
    ['general', 'General'], ['number', 'Number (1,234.50)'], ['integer', 'Whole number (1,235)'],
    ['currency', 'Currency ($1,234.50)'], ['percent', 'Percent (12%)'], ['percent1', 'Percent (12.3%)'],
    ['date', 'Short date (3/15/2026)'], ['longdate', 'Long date (March 15, 2026)'], ['text', 'Text (@)'],
  ];

  const TOOL_SETS = {
    basic: ['format', 'bold', 'fill', 'formulas', 'download', 'reset'],
    data: ['format', 'bold', 'fill', 'sort', 'textToColumns', 'removeDuplicates', 'import', 'csv', 'formulas', 'download', 'reset'],
    organize: ['format', 'bold', 'fill', 'sort', 'filter', 'table', 'validation', 'condFormat', 'names', 'formulas', 'download', 'reset'],
    analyze: ['format', 'bold', 'fill', 'sort', 'filter', 'table', 'condFormat', 'names', 'chart', 'pivot', 'formulas', 'download', 'reset'],
    all: ['format', 'bold', 'fill', 'sort', 'filter', 'textToColumns', 'removeDuplicates', 'import', 'csv', 'table', 'validation', 'condFormat', 'names', 'chart', 'pivot', 'query', 'sheets', 'formulas', 'download', 'reset'],
  };

  function mount(container, workbook, options = {}) {
    const opts = { tools: 'basic', highlights: [], fileName: 'workbook', ...options };
    const tools = new Set(Array.isArray(opts.tools) ? opts.tools : TOOL_SETS[opts.tools] || TOOL_SETS.basic);
    let wb = workbook;
    let sheetName = wb.sheets[0].name;
    let sel = { r1: 0, c1: 0, r2: 0, c2: 0, ar: 0, ac: 0 };
    let editing = null; // { value, origin: 'cell'|'bar', point: null|{start,end} }
    let dragging = false;
    let clip = null;
    let showFormulas = false;
    const filters = {}; // sheet -> { range, column, value }

    // ---------- DOM skeleton
    const root_ = h('div', { class: 'wb', 'data-testid': 'workbook' });
    const toolbar = h('div', { class: 'wb-toolbar', role: 'toolbar', 'aria-label': 'Workbook tools' });
    const nameBox = h('input', { class: 'wb-namebox', 'aria-label': 'Name Box', spellcheck: 'false' });
    const fx = h('input', { class: 'wb-fx', 'aria-label': 'Formula bar', spellcheck: 'false', autocomplete: 'off' });
    const fxHint = h('div', { class: 'wb-fx-hint', 'aria-live': 'polite' });
    const formulaRow = h('div', { class: 'wb-formula' }, nameBox, h('span', { class: 'wb-fx-label', 'aria-hidden': 'true' }, 'fx'), fx);
    const gridWrap = h('div', { class: 'wb-grid-wrap', tabindex: '0', 'aria-label': 'Spreadsheet grid. Use arrow keys to move and start typing to edit.' });
    const grid = h('table', { class: 'wb-grid', role: 'grid' });
    const editor = h('input', { class: 'wb-editor', 'aria-label': 'Cell editor', spellcheck: 'false', autocomplete: 'off' });
    gridWrap.append(grid, editor);
    const tabs = h('div', { class: 'wb-tabs', role: 'tablist' });
    const status = h('div', { class: 'wb-status', 'aria-live': 'polite' });
    const chartsArea = h('div', { class: 'wb-charts' });
    root_.append(toolbar, formulaRow, fxHint, gridWrap, h('div', { class: 'wb-bottom' }, tabs, status), chartsArea);
    container.innerHTML = '';
    container.append(root_);

    // ---------- helpers
    const sheet = () => wb.getSheet(sheetName) || wb.sheets[0];
    const addr = (r, c) => F.toAddress(r, c);
    const selRange = () => ({ r1: Math.min(sel.r1, sel.r2), c1: Math.min(sel.c1, sel.c2), r2: Math.max(sel.r1, sel.r2), c2: Math.max(sel.c1, sel.c2) });
    const selAddress = () => F.rangeToAddress(selRange());
    const isSingle = () => { const r = selRange(); return r.r1 === r.r2 && r.c1 === r.c2; };

    function dims() {
      const b = wb.usedBounds(sheetName);
      let rows = Math.max(b.rows + 6, 20), cols = Math.max(b.cols + 2, 8);
      wb.charts.forEach(() => {});
      rows = Math.max(rows, sel.r2 + 3, sel.ar + 3);
      cols = Math.max(cols, sel.c2 + 2, sel.ac + 2);
      return { rows: Math.min(rows, 500), cols: Math.min(cols, 40) };
    }

    function hiddenRows() {
      const f = filters[sheetName];
      if (!f) return new Set();
      const r = WB.normRange(f.range);
      const hidden = new Set();
      for (let i = r.r1 + 1; i <= r.r2; i++) {
        const v = wb.getValue(sheetName, { row: i, col: f.col });
        const text = F.formatValue(v, wb.getFormat(sheetName, { row: i, col: f.col }));
        if (!f.values.has(text)) hidden.add(i);
      }
      return hidden;
    }

    function highlightSet() {
      const m = new Map();
      (opts.highlights || []).forEach((hl, i) => {
        if ((hl.sheet || wb.sheets[0].name).toLowerCase() !== sheetName.toLowerCase()) return;
        const r = WB.normRange(hl.range);
        for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) m.set(row + ',' + col, hl.label || 'Task ' + (i + 1));
      });
      return m;
    }

    function refHighlights() {
      const input = editing ? editing.value : wb.getInput(sheetName, { row: sel.ar, col: sel.ac });
      if (!input || !input.startsWith('=')) return new Map();
      const m = new Map();
      F.referencedRanges(input).forEach((r, i) => {
        if (r.sheet && r.sheet.toLowerCase() !== sheetName.toLowerCase()) return;
        for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
          const k = row + ',' + col;
          const cls = ['ref' + (i % 6)];
          if (row === r.r1) cls.push('ref-t'); if (row === r.r2) cls.push('ref-b');
          if (col === r.c1) cls.push('ref-l'); if (col === r.c2) cls.push('ref-r');
          m.set(k, (m.get(k) || []).concat(cls));
        }
      });
      return m;
    }

    // ---------- rendering
    function render() {
      const { rows, cols } = dims();
      const sh = sheet();
      const hidden = hiddenRows();
      const hl = highlightSet();
      const refs = refHighlights();
      const r = selRange();
      const filter = filters[sheetName];
      const head = h('tr', {}, h('th', { class: 'wb-corner', 'aria-hidden': 'true' }));
      for (let c = 0; c < cols; c++) {
        const w = sh.widths && sh.widths[F.indexToCol(c)];
        head.append(h('th', { class: 'wb-colhead' + (c >= r.c1 && c <= r.c2 ? ' hdr-sel' : ''), style: w ? { width: w + 'px', minWidth: w + 'px' } : null, 'data-c': c, scope: 'col' }, F.indexToCol(c)));
      }
      const body = h('tbody');
      for (let row = 0; row < rows; row++) {
        if (hidden.has(row)) continue;
        const tr = h('tr', {}, h('th', { class: 'wb-rowhead' + (row >= r.r1 && row <= r.r2 ? ' hdr-sel' : ''), 'data-r': row, scope: 'row' }, String(row + 1)));
        for (let c = 0; c < cols; c++) {
          const key = addr(row, c);
          const cell = sh.cells[key];
          const v = cell ? wb.getValue(sheetName, { row, col: c }) : null;
          const fmt = cell ? wb.getFormat(sheetName, { row, col: c }) : null;
          let text = cell ? (showFormulas && String(cell.input).startsWith('=') ? String(cell.input) : F.formatValue(v, fmt)) : '';
          const cls = ['wb-cell'];
          if (typeof v === 'number' && !showFormulas) cls.push('num');
          if (typeof v === 'boolean') cls.push('bool');
          if (F.isError(v)) cls.push('err');
          if (cell && cell.locked) cls.push('locked');
          if (cell && cell.style && cell.style.bold) cls.push('bold');
          if (cell && cell.style && cell.style.align) cls.push('al-' + cell.style.align);
          if (cell && cell.style && cell.style.fill) cls.push('fill-' + cell.style.fill);
          if (row >= r.r1 && row <= r.r2 && c >= r.c1 && c <= r.c2) cls.push('sel');
          if (row === sel.ar && c === sel.ac) cls.push('active');
          if (hl.has(row + ',' + c)) cls.push('task');
          const rc = refs.get(row + ',' + c);
          if (rc) cls.push(...rc);
          const cf = cell ? wb.condStyle(sheetName, row, c) : null;
          if (cf) cls.push('cf-' + cf);
          const t = wb.tableAt(sheetName, row, c);
          if (t) {
            const tr0 = WB.normRange(t.ref);
            cls.push(row === tr0.r1 ? 'tbl-head' : (row - tr0.r1) % 2 ? 'tbl-odd' : 'tbl-even');
          }
          const w = sh.widths && sh.widths[F.indexToCol(c)];
          const td = h('td', {
            class: cls.join(' '), 'data-r': row, 'data-c': c, role: 'gridcell',
            title: F.isError(v) && v.hint ? v.code + ' — ' + v.hint : hl.get(row + ',' + c) || null,
            style: w ? { maxWidth: w + 'px' } : null,
          }, text);
          if (filter && row === WB.normRange(filter.range).r1 && c === filter.col) td.append(h('span', { class: 'filter-badge', title: 'Filtered' }, '▾'));
          tr.append(td);
        }
        body.append(tr);
      }
      grid.replaceChildren(h('thead', {}, head), body);
      renderTabs();
      renderStatus();
      renderCharts();
      syncBars();
      positionEditor();
      updateToolbarState();
    }

    function renderTabs() {
      tabs.replaceChildren(...wb.sheets.map((s) => {
        const b = h('button', {
          class: 'wb-tab' + (s.name === sheetName ? ' active' : ''), role: 'tab', 'aria-selected': s.name === sheetName ? 'true' : 'false', type: 'button',
          onClick: () => switchSheet(s.name),
          onDblclick: () => { if (tools.has('sheets')) renameSheetDialog(s.name); },
        }, s.name);
        return b;
      }), tools.has('sheets') ? h('button', { class: 'wb-tab wb-tab-add', type: 'button', title: 'New sheet', 'aria-label': 'Add sheet', onClick: () => { const s = wb.addSheet(); switchSheet(s.name); } }, '+') : '');
    }

    function renderStatus() {
      const r = selRange();
      const parts = [];
      if (!isSingle()) {
        const st = wb.stats(sheetName, r);
        if (st.numericCount) parts.push('Average: ' + F.formatGeneral(st.average), 'Count: ' + st.count, 'Sum: ' + F.formatGeneral(st.sum));
        else if (st.count) parts.push('Count: ' + st.count);
      }
      const f = filters[sheetName];
      if (f) parts.unshift('Filter on: showing ' + f.values.size + ' value(s)');
      status.textContent = parts.join('   ');
    }

    function renderCharts() {
      const charts = wb.charts.filter((c) => c.sheet === sheetName);
      chartsArea.replaceChildren(...charts.map((c) => {
        let svg;
        try { svg = root.DIM.Charts.render(wb.chartData(c), c); } catch (e) { svg = '<p class="muted">Chart error: ' + e.message + '</p>'; }
        return h('figure', { class: 'wb-chart', 'data-chart': c.id },
          h('div', { class: 'wb-chart-svg', html: svg }),
          h('figcaption', {},
            h('span', { class: 'muted' }, c.type + ' chart of ' + c.range),
            h('span', {},
              h('button', { class: 'btn btn-sm', type: 'button', onClick: () => chartDialog(c) }, 'Edit'),
              h('button', { class: 'btn btn-sm', type: 'button', onClick: () => wb.removeChart(c.id) }, 'Delete'))));
      }));
    }

    function syncBars() {
      nameBox.value = isSingle() ? addr(sel.ar, sel.ac) : selAddress();
      if (!editing || editing.origin !== 'bar') fx.value = editing ? editing.value : wb.getInput(sheetName, { row: sel.ar, col: sel.ac });
      const cell = wb.getCell(sheetName, { row: sel.ar, col: sel.ac });
      const v = cell ? wb.getValue(sheetName, { row: sel.ar, col: sel.ac }) : null;
      let hint = '';
      if (cell && cell.locked) hint = 'Locked: this cell holds given data.';
      if (F.isError(v)) hint = v.code + (v.hint ? ' — ' + v.hint : '');
      const perr = wb.formulaError(sheetName, { row: sel.ar, col: sel.ac });
      if (perr) hint = 'Formula problem: ' + perr;
      if (!hint && cell && typeof v === 'number') {
        const fmt = wb.getFormat(sheetName, { row: sel.ar, col: sel.ac });
        if (fmt === 'date' || fmt === 'longdate') hint = 'Stored as the number ' + v + ' (days since 1/1/1900), shown as a date.';
      }
      if (!hint && cell && typeof v === 'string' && F.parseNumberLike(v)) hint = 'This looks like a number but is stored as text.';
      fxHint.textContent = hint;
      fxHint.classList.toggle('show', !!hint);
    }

    function cellEl(r, c) { return grid.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]'); }

    function positionEditor() {
      if (!editing || editing.origin !== 'cell') { editor.style.display = 'none'; return; }
      const td = cellEl(sel.ar, sel.ac);
      if (!td) return;
      editor.style.display = 'block';
      editor.style.left = td.offsetLeft + 'px';
      editor.style.top = td.offsetTop + 'px';
      editor.style.minWidth = td.offsetWidth + 'px';
      editor.style.height = td.offsetHeight + 'px';
      if (editor.value !== editing.value) editor.value = editing.value;
    }

    function scrollIntoView() {
      const td = cellEl(sel.r2 === sel.ar && sel.c2 === sel.ac ? sel.ar : sel.r2, sel.r2 === sel.ar && sel.c2 === sel.ac ? sel.ac : sel.c2);
      if (!td) return;
      const wrap = gridWrap;
      const top = td.offsetTop, left = td.offsetLeft;
      const headH = 28, headW = 46;
      if (top - headH < wrap.scrollTop) wrap.scrollTop = top - headH;
      else if (top + td.offsetHeight > wrap.scrollTop + wrap.clientHeight) wrap.scrollTop = top + td.offsetHeight - wrap.clientHeight;
      if (left - headW < wrap.scrollLeft) wrap.scrollLeft = left - headW;
      else if (left + td.offsetWidth > wrap.scrollLeft + wrap.clientWidth) wrap.scrollLeft = left + td.offsetWidth - wrap.clientWidth;
    }

    // ---------- selection
    function setActive(r, c, extend) {
      r = Math.max(0, r); c = Math.max(0, c);
      if (extend) { sel.r2 = r; sel.c2 = c; }
      else sel = { r1: r, c1: c, r2: r, c2: c, ar: r, ac: c };
      render();
      scrollIntoView();
    }

    function switchSheet(name) {
      if (editing) commit();
      sheetName = name;
      sel = { r1: 0, c1: 0, r2: 0, c2: 0, ar: 0, ac: 0 };
      render();
    }

    function selectAddress(text) {
      let t = String(text).trim();
      if (!t) return;
      let target = F.parseRangeAddress(t);
      if (!target) {
        const nr = wb.resolveName(t);
        if (nr) target = { sheet: nr.sheet, r1: nr.r1, c1: nr.c1, r2: nr.r2, c2: nr.c2 };
        const tbl = wb.tables.find((x) => x.name.toLowerCase() === t.toLowerCase());
        if (tbl) target = { sheet: tbl.sheet, ...WB.normRange(tbl.ref) };
      }
      if (!target) { toast('"' + t + '" is not a cell address, range or name.', 'warn'); return; }
      if (target.sheet && wb.getSheet(target.sheet)) sheetName = wb.getSheet(target.sheet).name;
      sel = { r1: target.r1, c1: target.c1, r2: target.r2, c2: target.c2, ar: target.r1, ac: target.c1 };
      render();
      scrollIntoView();
      gridWrap.focus();
    }

    // ---------- editing
    function startEdit(initial, origin = 'cell') {
      const current = wb.getInput(sheetName, { row: sel.ar, col: sel.ac });
      if (wb.isLocked(sheetName, { row: sel.ar, col: sel.ac })) {
        toast('That cell is locked because it holds given data. Work in the empty cells.', 'warn');
        return;
      }
      editing = { value: initial !== undefined ? initial : current, origin, point: null };
      if (origin === 'cell') {
        editor.value = editing.value;
        positionEditor();
        editor.focus();
        const end = editor.value.length;
        editor.setSelectionRange(end, end);
      } else {
        fx.focus();
      }
      render();
    }

    function commit(moveR = 0, moveC = 0) {
      if (!editing) return true;
      const value = editing.value;
      const msg = wb.validate(sheetName, { row: sel.ar, col: sel.ac }, value);
      if (msg) {
        dialog('Data validation', h('p', {}, 'The value "' + value + '" is not allowed here. ' + msg));
        return false;
      }
      editing = null;
      editor.style.display = 'none';
      const before = wb.getInput(sheetName, { row: sel.ar, col: sel.ac });
      if (value !== before) {
        let input = value;
        // autocomplete missing closing parentheses, like Excel does
        if (input.startsWith('=')) {
          const open = (input.match(/\(/g) || []).length, close = (input.match(/\)/g) || []).length;
          if (open > close && !/"[^"]*$/.test(input)) input += ')'.repeat(open - close);
        }
        wb.setInput(sheetName, { row: sel.ar, col: sel.ac }, input);
      }
      if (moveR || moveC) {
        const r = selRange();
        if (!isSingle() && moveR) {
          // Enter moves within a multi-cell selection
          let nr = sel.ar + moveR, nc = sel.ac;
          if (nr > r.r2) { nr = r.r1; nc = sel.ac + 1 > r.c2 ? r.c1 : sel.ac + 1; }
          sel.ar = nr; sel.ac = nc;
        } else {
          sel = { r1: Math.max(0, sel.ar + moveR), c1: Math.max(0, sel.ac + moveC), r2: Math.max(0, sel.ar + moveR), c2: Math.max(0, sel.ac + moveC), ar: Math.max(0, sel.ar + moveR), ac: Math.max(0, sel.ac + moveC) };
        }
      }
      render();
      scrollIntoView();
      gridWrap.focus();
      return true;
    }

    function cancelEdit() {
      editing = null;
      editor.style.display = 'none';
      render();
      gridWrap.focus();
    }

    function pointable(text, caret) {
      if (!text.startsWith('=')) return false;
      const before = text.slice(0, caret).replace(/\s+$/, '');
      return /[=(,+\-*/^&<>:]$/.test(before);
    }

    // F4: cycle A1 -> $A$1 -> A$1 -> $A1 -> A1 for the reference at the caret
    function toggleAbsolute(input) {
      const text = input.value;
      const caret = input.selectionStart;
      const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
      let m;
      while ((m = re.exec(text))) {
        const start = m.index, end = m.index + m[0].length;
        if (caret >= start && caret <= end) {
          const prev = text[start - 1];
          if (prev && /[A-Za-z_"]/.test(prev)) continue;
          const state = (m[1] ? 2 : 0) + (m[3] ? 1 : 0); // 0 A1, 3 $A$1, 1 A$1, 2 $A1
          const nextState = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
          const rep = (nextState & 2 ? '$' : '') + m[2] + (nextState & 1 ? '$' : '') + m[4];
          input.value = text.slice(0, start) + rep + text.slice(end);
          input.setSelectionRange(start + rep.length, start + rep.length);
          return input.value;
        }
      }
      return text;
    }

    // ---------- clipboard
    function copySelection(cut) {
      const r = selRange();
      clip = wb.copy(sheetName, r);
      clip.cut = !!cut;
      const tsv = wb.getRangeValues(sheetName, r).map((row, i) => row.map((v, j) => F.formatValue(v, wb.getFormat(sheetName, { row: r.r1 + i, col: r.c1 + j }))).join('\t')).join('\n');
      clip.text = tsv;
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(tsv).catch(() => {});
      toast((cut ? 'Cut ' : 'Copied ') + F.rangeToAddress(r) + '. Select a cell and press Ctrl+V.');
    }

    function pasteFrom(text) {
      const target = { row: sel.ar, col: sel.ac };
      if (clip && (text === undefined || text === clip.text || text === '')) {
        wb.paste(sheetName, target, clip);
        if (clip.cut) {
          const src = { r1: clip.r1, c1: clip.c1, r2: clip.r1 + clip.cells.length - 1, c2: clip.c1 + clip.cells[0].length - 1 };
          if (clip.sheet !== sheetName || src.r1 !== target.row || src.c1 !== target.col) {
            for (let i = src.r1; i <= src.r2; i++) for (let j = src.c1; j <= src.c2; j++) {
              const inside = clip.sheet === sheetName && i >= target.row && i < target.row + clip.cells.length && j >= target.col && j < target.col + clip.cells[0].length;
              if (!inside) wb.setInput(clip.sheet, { row: i, col: j }, '');
            }
          }
          clip = null;
        }
      } else if (text) {
        wb.pasteText(sheetName, target, text);
      }
    }

    // ---------- events: grid
    gridWrap.addEventListener('mousedown', (e) => {
      const td = e.target.closest('td.wb-cell');
      const th = e.target.closest('th');
      if (th && th.classList.contains('wb-colhead')) {
        const c = +th.dataset.c;
        if (editing) commit();
        const b = dims();
        sel = { r1: 0, c1: c, r2: Math.max(wb.usedBounds(sheetName).rows - 1, 0), c2: c, ar: 0, ac: c };
        if (b) render();
        return;
      }
      if (th && th.classList.contains('wb-rowhead')) {
        const r = +th.dataset.r;
        if (editing) commit();
        sel = { r1: r, c1: 0, r2: r, c2: Math.max(wb.usedBounds(sheetName).cols - 1, 0), ar: r, ac: 0 };
        render();
        return;
      }
      if (!td) return;
      const r = +td.dataset.r, c = +td.dataset.c;
      if (editing) {
        const input = editing.origin === 'bar' ? fx : editor;
        const caret = input.selectionStart ?? editing.value.length;
        if (editing.point || pointable(editing.value, caret)) {
          e.preventDefault();
          const start = editing.point ? editing.point.start : caret;
          const end = editing.point ? editing.point.end : caret;
          const ref = addr(r, c);
          editing.value = editing.value.slice(0, start) + ref + editing.value.slice(end);
          editing.point = { start, end: start + ref.length, anchor: { r, c } };
          input.value = editing.value;
          dragging = 'point';
          render();
          input.focus();
          input.setSelectionRange(editing.point.end, editing.point.end);
          return;
        }
        if (!commit()) return;
      }
      e.preventDefault();
      gridWrap.focus();
      dragging = 'select';
      setActive(r, c, e.shiftKey);
    });

    gridWrap.addEventListener('mouseover', (e) => {
      if (!dragging) return;
      const td = e.target.closest('td.wb-cell');
      if (!td) return;
      const r = +td.dataset.r, c = +td.dataset.c;
      if (dragging === 'point' && editing && editing.point) {
        const a = editing.point.anchor;
        const ref = a.r === r && a.c === c ? addr(r, c) : F.rangeToAddress({ r1: Math.min(a.r, r), c1: Math.min(a.c, c), r2: Math.max(a.r, r), c2: Math.max(a.c, c) });
        editing.value = editing.value.slice(0, editing.point.start) + ref + editing.value.slice(editing.point.end);
        editing.point.end = editing.point.start + ref.length;
        const input = editing.origin === 'bar' ? fx : editor;
        input.value = editing.value;
        render();
        input.focus();
        input.setSelectionRange(editing.point.end, editing.point.end);
      } else if (dragging === 'select') {
        sel.r2 = r; sel.c2 = c;
        render();
      }
    });

    document.addEventListener('mouseup', onMouseUp);
    function onMouseUp() { dragging = false; }

    gridWrap.addEventListener('dblclick', (e) => {
      const td = e.target.closest('td.wb-cell');
      if (td) startEdit();
    });

    gridWrap.addEventListener('keydown', (e) => {
      if (e.target !== gridWrap) return;
      const ctrl = e.ctrlKey || e.metaKey;
      const k = e.key;
      const { rows, cols } = dims();
      const move = (dr, dc) => { e.preventDefault(); const nr = (e.shiftKey ? sel.r2 : sel.ar) + dr, nc = (e.shiftKey ? sel.c2 : sel.ac) + dc; setActive(Math.max(0, nr), Math.max(0, nc), e.shiftKey); };
      if (ctrl && k.startsWith('Arrow')) {
        e.preventDefault();
        const from = e.shiftKey ? { r: sel.r2, c: sel.c2 } : { r: sel.ar, c: sel.ac };
        const to = jumpEdge(from, k);
        setActive(to.r, to.c, e.shiftKey);
        return;
      }
      switch (k) {
        case 'ArrowUp': return move(-1, 0);
        case 'ArrowDown': return move(1, 0);
        case 'ArrowLeft': return move(0, -1);
        case 'ArrowRight': return move(0, 1);
        case 'Enter': e.preventDefault(); if (e.shiftKey) setActive(sel.ar - 1, sel.ac); else setActive(sel.ar + 1, sel.ac); return;
        case 'Tab': e.preventDefault(); setActive(sel.ar, sel.ac + (e.shiftKey ? -1 : 1)); return;
        case 'F2': e.preventDefault(); startEdit(); return;
        case 'Home': e.preventDefault(); setActive(ctrl ? 0 : sel.ar, 0); return;
        case 'End': if (ctrl) { e.preventDefault(); const b = wb.usedBounds(sheetName); setActive(Math.max(0, b.rows - 1), Math.max(0, b.cols - 1)); } return;
        case 'PageDown': e.preventDefault(); setActive(Math.min(rows + 10, sel.ar + 15), sel.ac); return;
        case 'PageUp': e.preventDefault(); setActive(Math.max(0, sel.ar - 15), sel.ac); return;
        case 'Delete':
        case 'Backspace': e.preventDefault(); wb.clearRange(sheetName, selRange()); return;
        case 'Escape': if (clip) { clip = null; toast('Copy cancelled.'); } return;
      }
      if (ctrl) {
        const lk = k.toLowerCase();
        if (lk === 'z') { e.preventDefault(); wb.undo(); return; }
        if (lk === 'y') { e.preventDefault(); wb.redo(); return; }
        if (lk === 'c') { e.preventDefault(); copySelection(false); return; }
        if (lk === 'x') { e.preventDefault(); copySelection(true); return; }
        if (lk === 'd') { e.preventDefault(); doFill('down'); return; }
        if (lk === 'r') { e.preventDefault(); doFill('right'); return; }
        if (lk === 'b') { e.preventDefault(); toggleBold(); return; }
        if (lk === 'a') { e.preventDefault(); const cr = wb.currentRegion(sheetName, { row: sel.ar, col: sel.ac }); sel = { ...cr, ar: sel.ar, ac: sel.ac }; render(); return; }
        if (k === '`') { e.preventDefault(); showFormulas = !showFormulas; render(); return; }
        return; // let Ctrl+V reach the paste event
      }
      if (k.length === 1 && !e.altKey) {
        e.preventDefault();
        startEdit(k);
      }
    });

    gridWrap.addEventListener('paste', (e) => {
      if (editing) return;
      e.preventDefault();
      const text = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
      pasteFrom(text);
    });

    function jumpEdge(from, key) {
      const dr = key === 'ArrowUp' ? -1 : key === 'ArrowDown' ? 1 : 0;
      const dc = key === 'ArrowLeft' ? -1 : key === 'ArrowRight' ? 1 : 0;
      const b = wb.usedBounds(sheetName);
      const filled = (r, c) => r >= 0 && c >= 0 && wb.getInput(sheetName, { row: r, col: c }) !== '';
      let r = from.r, c = from.c;
      const limitR = Math.max(b.rows, from.r + 1), limitC = Math.max(b.cols, from.c + 1);
      const inBounds = (rr, cc) => rr >= 0 && cc >= 0 && rr < limitR + 1 && cc < limitC + 1;
      if (filled(r, c) && filled(r + dr, c + dc)) {
        while (inBounds(r + dr, c + dc) && filled(r + dr, c + dc)) { r += dr; c += dc; }
      } else {
        r += dr; c += dc;
        while (inBounds(r + dr, c + dc) && !filled(r, c)) { r += dr; c += dc; }
        if (!filled(r, c)) { r = dr < 0 ? 0 : dr > 0 ? Math.max(0, limitR - 1) : r; c = dc < 0 ? 0 : dc > 0 ? Math.max(0, limitC - 1) : c; }
      }
      return { r: Math.max(0, r), c: Math.max(0, c) };
    }

    // ---------- events: editors
    function editorKeys(e, input) {
      if (!editing) return;
      if (e.key === 'Enter') { e.preventDefault(); commit(e.shiftKey ? -1 : 1, 0); }
      else if (e.key === 'Tab') { e.preventDefault(); commit(0, e.shiftKey ? -1 : 1); }
      else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      else if (e.key === 'F4') { e.preventDefault(); editing.value = toggleAbsolute(input); editing.point = null; render(); input.focus(); }
      else if (editing.origin === 'cell' && !editing.value.startsWith('=') && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && !e.shiftKey) {
        // "Enter mode": arrows commit and move, as in Excel
        e.preventDefault();
        const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
        commit(d[0], d[1]);
      }
    }
    editor.addEventListener('keydown', (e) => editorKeys(e, editor));
    editor.addEventListener('input', () => { if (editing) { editing.value = editor.value; editing.point = null; fx.value = editor.value; renderRefsOnly(); } });
    fx.addEventListener('focus', () => { if (!editing) startEdit(undefined, 'bar'); });
    fx.addEventListener('keydown', (e) => editorKeys(e, fx));
    fx.addEventListener('input', () => { if (!editing) startEdit(fx.value, 'bar'); editing.value = fx.value; editing.point = null; renderRefsOnly(); });
    nameBox.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); selectAddress(nameBox.value); } if (e.key === 'Escape') { syncBars(); gridWrap.focus(); } });
    nameBox.addEventListener('focus', () => nameBox.select());

    let refTimer = null;
    function renderRefsOnly() {
      clearTimeout(refTimer);
      refTimer = setTimeout(() => { const keepFocus = document.activeElement; const s = keepFocus && keepFocus.selectionStart; render(); if (keepFocus && keepFocus.focus) { keepFocus.focus(); if (s !== null && s !== undefined && keepFocus.setSelectionRange) keepFocus.setSelectionRange(s, s); } }, 120);
    }

    // ---------- toolbar
    function btn(label, title, onClick, extra = {}) {
      return h('button', { class: 'tb-btn', type: 'button', title, 'aria-label': title, onClick: () => { if (editing) commit(); onClick(); }, ...extra }, label);
    }

    const undoBtn = btn('↶', 'Undo (Ctrl+Z)', () => wb.undo());
    const redoBtn = btn('↷', 'Redo (Ctrl+Y)', () => wb.redo());
    const formatSel = select(FORMAT_OPTIONS, 'general', { class: 'tb-select', 'aria-label': 'Number format', title: 'Number format' });
    formatSel.addEventListener('change', () => { wb.setFormat(sheetName, selRange(), formatSel.value); gridWrap.focus(); });

    function group(label, ...items) { return h('div', { class: 'tb-group', role: 'group', 'aria-label': label }, ...items, h('span', { class: 'tb-group-label' }, label)); }

    function buildToolbar() {
      const groups = [group('Edit', undoBtn, redoBtn)];
      const cellsG = [];
      if (tools.has('format')) cellsG.push(formatSel);
      if (tools.has('bold')) cellsG.push(btn('B', 'Bold (Ctrl+B)', toggleBold, { class: 'tb-btn tb-bold' }));
      if (tools.has('fill')) cellsG.push(btn('Fill ↓', 'Fill Down (Ctrl+D): copy the top cell of the selection down', () => doFill('down')), btn('Fill →', 'Fill Right (Ctrl+R)', () => doFill('right')));
      if (cellsG.length) groups.push(group('Cells', ...cellsG));
      const dataG = [];
      if (tools.has('sort')) dataG.push(btn('Sort', 'Sort the data around the selected cell', sortDialog));
      if (tools.has('filter')) dataG.push(btn('Filter', 'Filter rows by a column value', filterDialog));
      if (tools.has('textToColumns')) dataG.push(btn('Text to Columns', 'Split one column into several using a delimiter', textToColumnsDialog));
      if (tools.has('removeDuplicates')) dataG.push(btn('Remove Duplicates', 'Remove duplicate rows', removeDuplicatesDialog));
      if (tools.has('validation')) dataG.push(btn('Data Validation', 'Only allow certain values in the selected cells', validationDialog));
      if (tools.has('import')) dataG.push(btn('From Text/CSV', 'Import a delimited text file into a new sheet', importDialog));
      if (tools.has('query')) dataG.push(btn('Get Data (Power Query)', 'Open the Power Query editor', () => root.DIM.QueryEditor.open(wb, { onLoad: (name) => switchSheet(name) })));
      if (tools.has('query') || tools.has('pivot')) dataG.push(btn('Refresh All', 'Re-run every query and refresh every PivotTable', () => { const res = wb.refreshAll(); const bad = res.find((r) => r && r.error); toast(bad ? 'Refresh problem: ' + bad.error : 'Refreshed.', bad ? 'warn' : 'ok'); }));
      if (dataG.length) groups.push(group('Data', ...dataG));
      const insG = [];
      if (tools.has('table')) insG.push(btn('Format as Table', 'Turn the data into an Excel Table (Ctrl+T)', tableDialog));
      if (tools.has('condFormat')) insG.push(btn('Highlight Rules', 'Conditional formatting', condFormatDialog));
      if (tools.has('names')) insG.push(btn('Define Name', 'Give the selected range a name', nameDialog));
      if (tools.has('chart')) insG.push(btn('Chart', 'Insert a chart from the selected data', () => chartDialog()));
      if (tools.has('pivot')) insG.push(btn('PivotTable', 'Summarize data with a PivotTable', pivotDialog));
      if (insG.length) groups.push(group('Insert', ...insG));
      const fileG = [];
      if (tools.has('formulas')) fileG.push(btn('Show Formulas', 'Show formulas instead of results (Ctrl+`)', () => { showFormulas = !showFormulas; render(); }, { 'aria-pressed': 'false', class: 'tb-btn tb-toggle-formulas' }));
      if (tools.has('csv')) fileG.push(btn('Save as CSV', 'Download this sheet as a .csv file', () => downloadText(WB.sheetToCSV(wb, sheetName), safeName(opts.fileName + '-' + sheetName) + '.csv', 'text/csv')));
      if (tools.has('download')) fileG.push(btn('Download .xlsx', 'Download this workbook to open in Excel', () => root.DIM.Xlsx.download(root.DIM.Xlsx.write(wb, { title: opts.fileName }), safeName(opts.fileName) + '.xlsx')));
      if (tools.has('reset') && opts.onReset) fileG.push(btn('Reset', 'Start this workbook over', () => {
        dialog('Reset workbook?', h('p', {}, 'This erases your work in this workbook and restores the original data.'), [
          { label: 'Cancel' },
          { label: 'Reset', primary: true, onClick: () => { setWorkbook(opts.onReset()); toast('Workbook reset.'); } },
        ]);
      }));
      if (fileG.length) groups.push(group('File', ...fileG));
      toolbar.replaceChildren(...groups);
    }

    function updateToolbarState() {
      undoBtn.disabled = !wb.canUndo;
      redoBtn.disabled = !wb.canRedo;
      const fmt = wb.getFormat(sheetName, { row: sel.ar, col: sel.ac }) || 'general';
      formatSel.value = FORMAT_OPTIONS.some((o) => o[0] === fmt) ? fmt : 'general';
      const tf = toolbar.querySelector('.tb-toggle-formulas');
      if (tf) { tf.setAttribute('aria-pressed', String(showFormulas)); tf.classList.toggle('on', showFormulas); }
    }

    function safeName(s) { return String(s).replace(/[^\w.-]+/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '') || 'workbook'; }

    function toggleBold() {
      const c = wb.getCell(sheetName, { row: sel.ar, col: sel.ac });
      const bold = !(c && c.style && c.style.bold);
      wb.setStyle(sheetName, selRange(), { bold });
    }

    function doFill(dir) {
      const r = selRange();
      if ((dir === 'down' && r.r1 === r.r2) || (dir === 'right' && r.c1 === r.c2)) {
        toast('Select the cell with the formula and the cells to fill (e.g. drag from D2 down to D11), then Fill.', 'warn');
        return;
      }
      wb.fill(sheetName, r, dir);
    }

    // The data block the user means: the selection if it is a block, else the current region.
    function dataBlock() {
      const r = selRange();
      if (r.r2 > r.r1 && r.c2 >= r.c1) return r;
      return wb.currentRegion(sheetName, { row: sel.ar, col: sel.ac });
    }

    function headerNames(r) {
      const out = [];
      for (let c = r.c1; c <= r.c2; c++) {
        const v = wb.getValue(sheetName, { row: r.r1, col: c });
        out.push([c - r.c1, (v === null ? 'Column ' + F.indexToCol(c) : F.toText(v)) + ' (' + F.indexToCol(c) + ')']);
      }
      return out;
    }

    // ---------- dialogs
    function sortDialog() {
      const r = dataBlock();
      const hdr = h('input', { type: 'checkbox', checked: true });
      const colSel = select(headerNames(r), sel.ac - r.c1);
      const orderSel = select([['asc', 'A → Z / smallest to largest'], ['desc', 'Z → A / largest to smallest']], 'asc');
      dialog('Sort ' + F.rangeToAddress(r), h('div', {},
        field('Sort by', colSel), field('Order', orderSel),
        h('label', { class: 'check' }, hdr, ' My data has headers'),
        h('p', { class: 'muted' }, 'Sorting moves whole rows so every record stays together.')), [
        { label: 'Cancel' },
        { label: 'Sort', primary: true, onClick: () => {
          const res = wb.sortRange(sheetName, r, Number(colSel.value), orderSel.value === 'asc', hdr.checked);
          if (res && res.error) { toast(res.error, 'warn'); return false; }
          sel = { r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2, ar: r.r1, ac: r.c1 };
          render();
        } },
      ]);
    }

    function filterDialog() {
      if (filters[sheetName]) { delete filters[sheetName]; render(); toast('Filter cleared.'); return; }
      const r = dataBlock();
      const colSel = select(headerNames(r), sel.ac - r.c1);
      const list = h('div', { class: 'checklist' });
      const fill = () => {
        const c = r.c1 + Number(colSel.value);
        const vals = new Set();
        for (let i = r.r1 + 1; i <= r.r2; i++) vals.add(F.formatValue(wb.getValue(sheetName, { row: i, col: c }), wb.getFormat(sheetName, { row: i, col: c })));
        list.replaceChildren(...[...vals].sort().map((v) => h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: true, value: v }), ' ' + (v === '' ? '(Blanks)' : v))));
      };
      colSel.addEventListener('change', fill);
      fill();
      dialog('Filter ' + F.rangeToAddress(r), h('div', {}, field('Column', colSel), h('p', { class: 'muted' }, 'Show rows where the value is:'), list,
        h('p', { class: 'muted' }, 'Filtering hides rows; it does not delete them. Click Filter again to clear it.')), [
        { label: 'Cancel' },
        { label: 'Apply', primary: true, onClick: () => {
          const values = new Set([...list.querySelectorAll('input:checked')].map((x) => x.value));
          filters[sheetName] = { range: F.rangeToAddress(r), col: r.c1 + Number(colSel.value), values };
          render();
        } },
      ]);
    }

    function textToColumnsDialog() {
      const r = selRange();
      if (r.c1 !== r.c2) { toast('Select cells in a single column first.', 'warn'); return; }
      const delimSel = select([[',', 'Comma ,'], ['tab', 'Tab'], [';', 'Semicolon ;'], ['|', 'Pipe |'], ['space', 'Space'], ['-', 'Dash -'], ['other', 'Other…']], ',');
      const other = h('input', { type: 'text', maxlength: '3', placeholder: 'e.g. /', style: { width: '5em' } });
      const preview = h('div', { class: 'preview-table' });
      const textCols = h('div', { class: 'checklist inline' });
      const delim = () => (delimSel.value === 'other' ? other.value : delimSel.value);
      const update = () => {
        const d = delim() === 'tab' ? '\t' : delim() === 'space' ? ' ' : delim();
        const rows = [];
        for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 4); i++) {
          const v = wb.getValue(sheetName, { row: i, col: r.c1 });
          rows.push(d ? WB.splitDelimited(F.toText(v ?? ''), d).map((x) => x.trim()) : [F.toText(v ?? '')]);
        }
        const n = Math.max(...rows.map((x) => x.length));
        preview.replaceChildren(h('table', {}, rows.map((row) => h('tr', {}, Array.from({ length: n }, (_, j) => h('td', {}, row[j] ?? ''))))));
        textCols.replaceChildren(...Array.from({ length: n }, (_, j) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: j }), ' Column ' + (j + 1) + ' as text')));
      };
      delimSel.addEventListener('change', update);
      other.addEventListener('input', update);
      update();
      dialog('Text to Columns', h('div', {},
        h('p', { class: 'muted' }, 'Splits ' + F.rangeToAddress(r) + ' into the columns to the right. Anything already there is overwritten.'),
        field('Delimiter', h('span', { class: 'row' }, delimSel, other)),
        h('p', { class: 'field-label' }, 'Preview'), preview,
        h('p', { class: 'field-label' }, 'Keep as text (protects leading zeros like 02134)'), textCols), [
        { label: 'Cancel' },
        { label: 'Finish', primary: true, onClick: () => {
          const d = delim();
          if (!d) { toast('Choose a delimiter.', 'warn'); return false; }
          const tc = [...textCols.querySelectorAll('input:checked')].map((x) => Number(x.value));
          const res = wb.textToColumns(sheetName, r, d, { textColumns: tc });
          if (res && res.error) { toast(res.error, 'warn'); return false; }
        } },
      ], { wide: true });
    }

    function removeDuplicatesDialog() {
      const r = dataBlock();
      const hdr = h('input', { type: 'checkbox', checked: true });
      const cols = h('div', { class: 'checklist' }, headerNames(r).map(([i, label]) => h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: true, value: i }), ' ' + label)));
      dialog('Remove Duplicates in ' + F.rangeToAddress(r), h('div', {}, h('label', { class: 'check' }, hdr, ' My data has headers'), h('p', { class: 'muted' }, 'A row is a duplicate when the checked columns all match another row.'), cols), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          const which = [...cols.querySelectorAll('input:checked')].map((x) => Number(x.value));
          if (!which.length) { toast('Pick at least one column.', 'warn'); return false; }
          const res = wb.removeDuplicates(sheetName, r, hdr.checked, which);
          dialog('Remove Duplicates', h('p', {}, res.removed ? res.removed + ' duplicate row(s) found and removed; ' + res.remaining + ' unique row(s) remain.' : 'No duplicate values found.'));
        } },
      ]);
    }

    function validationDialog() {
      const r = selRange();
      const typeSel = select([['list', 'List of allowed values'], ['whole', 'Whole number between'], ['decimal', 'Decimal number between'], ['textLength', 'Text length between']], 'list');
      const values = h('input', { type: 'text', placeholder: 'e.g. Freshman, Sophomore, Junior, Senior' });
      const min = h('input', { type: 'number', placeholder: 'minimum' });
      const max = h('input', { type: 'number', placeholder: 'maximum' });
      const msg = h('input', { type: 'text', placeholder: 'Optional error message' });
      const listRow = field('Allowed values (comma-separated)', values);
      const numRow = h('div', { class: 'row' }, field('Minimum', min), field('Maximum', max));
      const upd = () => { listRow.style.display = typeSel.value === 'list' ? '' : 'none'; numRow.style.display = typeSel.value === 'list' ? 'none' : ''; };
      typeSel.addEventListener('change', upd); upd();
      dialog('Data Validation for ' + F.rangeToAddress(r), h('div', {}, field('Allow', typeSel), listRow, numRow, field('Error message', msg)), [
        { label: 'Cancel' },
        { label: 'Apply', primary: true, onClick: () => {
          const rule = { type: typeSel.value };
          if (rule.type === 'list') {
            rule.values = values.value.split(',').map((x) => x.trim()).filter(Boolean);
            if (!rule.values.length) { toast('Type at least one allowed value.', 'warn'); return false; }
          } else {
            if (min.value !== '') rule.min = Number(min.value);
            if (max.value !== '') rule.max = Number(max.value);
          }
          if (msg.value) rule.message = msg.value;
          wb.addValidation(sheetName, F.rangeToAddress(r), rule);
          toast('Validation added to ' + F.rangeToAddress(r) + '.', 'ok');
        } },
      ]);
    }

    function condFormatDialog() {
      const r = isSingle() ? dataBlock() : selRange();
      const opSel = select([['>', 'Greater than'], ['<', 'Less than'], ['>=', 'Greater than or equal to'], ['<=', 'Less than or equal to'], ['between', 'Between'], ['=', 'Equal to'], ['contains', 'Text that contains'], ['duplicate', 'Duplicate values'], ['top', 'Top N values'], ['aboveAverage', 'Above average'], ['blank', 'Blank cells']], '>');
      const v1 = h('input', { type: 'text', placeholder: 'value' });
      const v2 = h('input', { type: 'text', placeholder: 'and' });
      const styleSel = select([['green', 'Green fill'], ['red', 'Light red fill'], ['yellow', 'Yellow fill']], 'green');
      const rangeIn = h('input', { type: 'text', value: F.rangeToAddress(r) });
      const upd = () => { v1.style.display = ['duplicate', 'aboveAverage', 'blank'].includes(opSel.value) ? 'none' : ''; v2.style.display = opSel.value === 'between' ? '' : 'none'; };
      opSel.addEventListener('change', upd); upd();
      dialog('Highlight Cells Rules', h('div', {}, field('Apply to', rangeIn), field('Format cells that are', opSel), h('div', { class: 'row' }, v1, v2), field('With', styleSel)), [
        { label: 'Clear all rules on this sheet', onClick: () => { wb.clearCondFormats(sheetName); } },
        { label: 'Cancel' },
        { label: 'Apply', primary: true, onClick: () => {
          if (!F.parseRangeAddress(rangeIn.value)) { toast('Enter a valid range like B2:B20.', 'warn'); return false; }
          const num = (x) => (x !== '' && !Number.isNaN(Number(x)) ? Number(x) : x);
          wb.addCondFormat(sheetName, rangeIn.value.toUpperCase(), { op: opSel.value, value: num(v1.value), value2: num(v2.value), style: styleSel.value });
        } },
      ]);
    }

    function tableDialog() {
      const r = dataBlock();
      const rangeIn = h('input', { type: 'text', value: F.rangeToAddress(r) });
      const nameIn = h('input', { type: 'text', value: 'Table' + (wb.tables.length + 1) });
      dialog('Format as Table', h('div', {}, field('Where is the data for your table?', rangeIn), field('Table name', nameIn, 'Use it in formulas: =SUM(' + nameIn.value + '[Sales])'), h('p', { class: 'muted' }, 'The first row becomes the header row. Typing in the row right under the table makes it grow automatically.')), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          const rr = F.parseRangeAddress(rangeIn.value);
          if (!rr) { toast('Enter a valid range.', 'warn'); return false; }
          const res = wb.addTable(sheetName, rr, nameIn.value.trim());
          if (res.error) { toast(res.error, 'warn'); return false; }
          toast('Created table ' + res.name + '.', 'ok');
        } },
      ]);
    }

    function nameDialog() {
      const r = selRange();
      const nameIn = h('input', { type: 'text', placeholder: 'e.g. TaxRate' });
      const refIn = h('input', { type: 'text', value: F.quoteSheet(sheetName) + '!' + F.rangeToAddress({ r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2') });
      const list = h('ul', { class: 'muted' }, Object.entries(wb.names).map(([n, ref]) => h('li', {}, n + ' → ' + ref)));
      dialog('Define Name', h('div', {}, field('Name', nameIn), field('Refers to', refIn), Object.keys(wb.names).length ? h('div', {}, h('p', { class: 'field-label' }, 'Existing names'), list) : null), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          if (!F.parseRangeAddress(refIn.value)) { toast('Refers to must be a range like Sheet1!$A$1:$A$10.', 'warn'); return false; }
          const res = wb.defineName(nameIn.value.trim(), refIn.value.trim());
          if (res.error) { toast(res.error, 'warn'); return false; }
          toast('Name ' + res.name + ' defined.', 'ok');
        } },
      ]);
    }

    function chartDialog(existing) {
      const r = existing ? WB.normRange(existing.range) : dataBlock();
      const typeSel = select([['column', 'Column (compare categories)'], ['bar', 'Bar (long category names)'], ['line', 'Line (change over time)'], ['pie', 'Pie (parts of one whole)'], ['scatter', 'Scatter (relationship between two numbers)']], existing ? existing.type : 'column');
      const titleIn = h('input', { type: 'text', value: existing ? existing.title : '', placeholder: 'Chart title' });
      const rangeIn = h('input', { type: 'text', value: existing ? existing.range : F.rangeToAddress(r) });
      const axisIn = h('input', { type: 'number', value: existing && existing.axisMin !== null && existing.axisMin !== undefined ? existing.axisMin : '', placeholder: 'auto (0)' });
      const prev = h('div', { class: 'chart-preview' });
      const draw = () => {
        const rr = F.parseRangeAddress(rangeIn.value);
        if (!rr) { prev.textContent = 'Enter a valid range.'; return; }
        const c = { type: typeSel.value, title: titleIn.value, range: F.rangeToAddress(rr), sheet: sheetName, axisMin: axisIn.value === '' ? null : Number(axisIn.value) };
        prev.innerHTML = root.DIM.Charts.render(wb.chartData(c), c);
      };
      [typeSel, titleIn, rangeIn, axisIn].forEach((x) => x.addEventListener('input', draw));
      draw();
      dialog(existing ? 'Edit chart' : 'Insert chart', h('div', { class: 'chart-dialog' },
        h('div', {}, field('Chart type', typeSel), field('Data range', rangeIn, 'First column = labels, first row = series names.'), field('Title', titleIn), field('Value axis starts at', axisIn, 'Leave blank to start at 0. Starting higher exaggerates differences.')),
        prev), [
        { label: 'Cancel' },
        { label: existing ? 'Save' : 'Insert', primary: true, onClick: () => {
          const rr = F.parseRangeAddress(rangeIn.value);
          if (!rr) { toast('Enter a valid range.', 'warn'); return false; }
          const def = { type: typeSel.value, title: titleIn.value, range: F.rangeToAddress(rr), axisMin: axisIn.value === '' ? null : Number(axisIn.value) };
          if (existing) wb.updateChart(existing.id, def); else wb.addChart({ sheet: sheetName, ...def });
        } },
      ], { wide: true });
    }

    function pivotDialog() {
      const r = dataBlock();
      const rangeIn = h('input', { type: 'text', value: F.rangeToAddress(r) });
      const recs = () => wb.recordsFrom(sheetName, rangeIn.value);
      let rec = recs();
      if (rec.error) { toast(rec.error, 'warn'); return; }
      const fieldOpts = () => rec.columns.map((c) => [c, c]);
      const rowsSel = select(fieldOpts(), rec.columns[0]);
      const colsSel = select([['', '(none)'], ...fieldOpts()], '');
      const numericCol = rec.columns.find((c) => rec.rows.some((x) => typeof x[c] === 'number')) || rec.columns[rec.columns.length - 1];
      const valsSel = select(fieldOpts(), numericCol);
      const aggSel = select([['sum', 'Sum'], ['count', 'Count'], ['average', 'Average'], ['min', 'Min'], ['max', 'Max']], 'sum');
      const groupSel = select([['', 'No grouping'], ['month', 'Months'], ['quarter', 'Quarters'], ['year', 'Years']], '');
      const prev = h('div', { class: 'preview-table' });
      const draw = () => {
        rec = recs();
        if (rec.error) { prev.textContent = rec.error; return; }
        const out = root.DIM.Pivot.build(rec.rows, { rows: rowsSel.value, columns: colsSel.value || null, values: valsSel.value, agg: aggSel.value, rowsGroup: groupSel.value || null });
        if (out.error) { prev.textContent = out.error; return; }
        prev.replaceChildren(h('table', {}, out.matrix.slice(0, 12).map((row, i) => h('tr', {}, row.map((v) => h(i < out.headerRows ? 'th' : 'td', {}, typeof v === 'number' ? F.formatGeneral(v) : v ?? ''))))));
      };
      [rowsSel, colsSel, valsSel, aggSel, groupSel].forEach((x) => x.addEventListener('change', draw));
      rangeIn.addEventListener('change', draw);
      draw();
      dialog('Create PivotTable', h('div', { class: 'chart-dialog' },
        h('div', {}, field('Source range', rangeIn), field('Rows', rowsSel), field('Group row dates by', groupSel), field('Columns (optional)', colsSel), field('Values', valsSel), field('Summarize values by', aggSel)),
        h('div', {}, h('p', { class: 'field-label' }, 'Preview'), prev)), [
        { label: 'Cancel' },
        { label: 'Create on new sheet', primary: true, onClick: () => {
          const res = wb.addPivot({ sourceSheet: sheetName, sourceRange: rangeIn.value, rows: rowsSel.value, columns: colsSel.value || null, values: valsSel.value, agg: aggSel.value, rowsGroup: groupSel.value || null });
          if (res.error) { toast(res.error, 'warn'); return false; }
          switchSheet(res.targetSheet);
          toast('PivotTable created on ' + res.targetSheet + '. It updates when you click Refresh All.', 'ok');
        } },
      ], { wide: true });
    }

    function importDialog() {
      const sources = Object.keys(wb.sources || {});
      const srcSel = select([['', sources.length ? 'Choose a provided file…' : 'No provided files'], ...sources.map((s) => [s, s])], '');
      const fileIn = h('input', { type: 'file', accept: '.csv,.txt,.tsv,text/plain,text/csv' });
      const delimSel = select([['auto', 'Detect automatically'], [',', 'Comma'], ['tab', 'Tab'], [';', 'Semicolon'], ['|', 'Pipe']], 'auto');
      const raw = h('pre', { class: 'raw-preview' });
      const prev = h('div', { class: 'preview-table' });
      const textCols = h('div', { class: 'checklist inline' });
      let text = '', fname = 'Import';
      const draw = () => {
        raw.textContent = text.split(/\r?\n/).slice(0, 6).join('\n') || '(choose a file to see its raw text)';
        if (!text) { prev.replaceChildren(); textCols.replaceChildren(); return; }
        const d = delimSel.value === 'auto' ? WB.detectDelimiter(text) : delimSel.value === 'tab' ? '\t' : delimSel.value;
        const rows = WB.parseDelimited(text, d).slice(0, 6);
        prev.replaceChildren(h('table', {}, rows.map((row, i) => h('tr', {}, row.map((v) => h(i ? 'td' : 'th', {}, v))))));
        textCols.replaceChildren(...(rows[0] || []).map((hname, j) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: j }), ' ' + (hname || 'Column ' + (j + 1)) + ' as text')));
      };
      srcSel.addEventListener('change', () => { if (srcSel.value) { text = wb.sources[srcSel.value]; fname = srcSel.value.replace(/\.\w+$/, ''); draw(); } });
      fileIn.addEventListener('change', async () => { const f = fileIn.files[0]; if (!f) return; text = await readFile(f); fname = f.name.replace(/\.\w+$/, ''); draw(); });
      delimSel.addEventListener('change', draw);
      draw();
      dialog('Import from Text/CSV', h('div', {},
        sources.length ? field('Provided file', srcSel) : null, field('Or upload a file', fileIn),
        h('p', { class: 'field-label' }, 'Raw file (this is all a CSV really is: text and delimiters)'), raw,
        field('Delimiter', delimSel), h('p', { class: 'field-label' }, 'Preview'), prev,
        h('p', { class: 'field-label' }, 'Load these columns as text (keeps leading zeros)'), textCols), [
        { label: 'Cancel' },
        { label: 'Load', primary: true, onClick: () => {
          if (!text) { toast('Choose a file first.', 'warn'); return false; }
          const tc = [...textCols.querySelectorAll('input:checked')].map((x) => Number(x.value));
          const s = wb.importDelimited(text, { delimiter: delimSel.value === 'auto' ? undefined : delimSel.value, textColumns: tc, sheetName: fname.slice(0, 28) });
          switchSheet(s.name);
          toast('Loaded ' + fname + ' into sheet ' + s.name + '.', 'ok');
        } },
      ], { wide: true });
    }

    function renameSheetDialog(name) {
      const input = h('input', { type: 'text', value: name });
      dialog('Rename sheet', field('New name', input), [
        { label: 'Delete sheet', onClick: () => { if (!wb.removeSheet(name)) { toast('A workbook needs at least one sheet.', 'warn'); return false; } switchSheet(wb.sheets[0].name); } },
        { label: 'Cancel' },
        { label: 'Rename', primary: true, onClick: () => { if (!wb.renameSheet(name, input.value.trim())) { toast('That name is empty or already used.', 'warn'); return false; } sheetName = input.value.trim(); render(); } },
      ]);
    }

    // ---------- wiring
    let unsub = wb.onChange(() => { render(); if (opts.onChange) opts.onChange(wb); });

    function setWorkbook(next) {
      unsub();
      wb = next;
      if (!wb.getSheet(sheetName)) sheetName = wb.sheets[0].name;
      sel = { r1: 0, c1: 0, r2: 0, c2: 0, ar: 0, ac: 0 };
      editing = null;
      unsub = wb.onChange(() => { render(); if (opts.onChange) opts.onChange(wb); });
      render();
      if (opts.onChange) opts.onChange(wb);
    }

    buildToolbar();
    render();

    return {
      get workbook() { return wb; },
      get sheetName() { return sheetName; },
      setWorkbook,
      switchSheet,
      select: selectAddress,
      commit: () => commit(),
      render,
      setHighlights(hl) { opts.highlights = hl; render(); },
      destroy() { unsub(); document.removeEventListener('mouseup', onMouseUp); container.innerHTML = ''; },
    };
  }

  root.DIM = root.DIM || {};
  root.DIM.SheetView = { mount, TOOL_SETS, FORMAT_OPTIONS };
})(typeof globalThis !== 'undefined' ? globalThis : this);
