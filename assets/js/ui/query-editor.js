/*
 * query-editor.js — the Power Query Editor window.
 *
 * Layout mirrors Excel's: a ribbon of transforms, a data preview, and a Query
 * Settings pane with the Applied Steps list. Clicking a step previews the data
 * at that point; new steps are inserted after the selected step. "Close & Load"
 * writes the result to a sheet named after the query; "Refresh All" on the
 * workbook re-runs every step against the current source files.
 */
(function (root) {
  'use strict';

  const F = root.DIM.Formula;
  const Q = root.DIM.Query;
  const { h, dialog, toast, field, select } = root.DIM.Dom;

  const TYPE_ICON = { text: 'ABC', number: '1.2', whole: '123', currency: '$', date: '📅', boolean: '✓✗', any: 'ABC123' };

  function open(wb, opts = {}) {
    let def = null;
    let stepIndex = 0; // number of steps applied in the preview
    let selectedCol = null;

    const queriesList = h('div', { class: 'pq-queries' });
    const ribbon = h('div', { class: 'pq-ribbon', role: 'toolbar', 'aria-label': 'Power Query transforms' });
    const preview = h('div', { class: 'pq-preview' });
    const stepsList = h('ol', { class: 'pq-steps', 'aria-label': 'Applied steps' });
    const nameIn = h('input', { type: 'text', 'aria-label': 'Query name' });
    const msg = h('div', { class: 'pq-msg', 'aria-live': 'polite' });
    const body = h('div', { class: 'pq' },
      h('aside', { class: 'pq-left' }, h('h4', {}, 'Queries'), queriesList, h('button', { class: 'btn btn-sm', type: 'button', onClick: newQueryDialog }, '+ New query')),
      h('section', { class: 'pq-main' }, ribbon, msg, preview),
      h('aside', { class: 'pq-right' }, h('h4', {}, 'Query Settings'), field('Name', nameIn), h('h4', {}, 'Applied Steps'), stepsList));
    nameIn.addEventListener('change', () => {
      if (!def) return;
      const n = nameIn.value.trim().replace(/[^\w ]/g, '').slice(0, 28);
      if (!n) { nameIn.value = def.name; return; }
      if (wb.queries.some((q) => q.id !== def.id && q.name.toLowerCase() === n.toLowerCase())) { toast('Another query already uses that name.', 'warn'); nameIn.value = def.name; return; }
      def.name = n;
      save();
    });

    const d = dialog('Power Query Editor', body, [
      { label: 'Close' },
      { label: 'Close & Load', primary: true, onClick: () => {
        if (!def) { toast('Create a query first.', 'warn'); return false; }
        save();
        const res = wb.loadQuery(def.id);
        if (res.error) { toast('Cannot load: ' + res.error, 'warn'); return false; }
        toast('Loaded ' + res.table.rows.length + ' rows to sheet "' + def.name + '".', 'ok');
        if (opts.onLoad) opts.onLoad(def.name);
      } },
    ], { wide: true });
    d.box.classList.add('modal-pq');

    function save() { if (def) wb.saveQuery(def); renderAll(); }

    function current() {
      if (!def) return null;
      return wb.runQuery({ ...def, steps: def.steps.slice(0, stepIndex) });
    }

    function renderAll() {
      renderQueries();
      renderSteps();
      renderPreview();
      renderRibbon();
    }

    function renderQueries() {
      queriesList.replaceChildren(...wb.queries.map((q) => h('button', {
        class: 'pq-query' + (def && q.id === def.id ? ' active' : ''), type: 'button',
        onClick: () => { def = JSON.parse(JSON.stringify(q)); stepIndex = def.steps.length; selectedCol = null; nameIn.value = def.name; renderAll(); },
      }, q.name, h('span', { class: 'muted' }, ' · ' + q.steps.length + ' steps'))));
      if (!wb.queries.length) queriesList.append(h('p', { class: 'muted' }, 'No queries yet.'));
    }

    function renderSteps() {
      if (!def) { stepsList.replaceChildren(); return; }
      const items = [h('li', { class: 'pq-step' + (stepIndex === 0 ? ' active' : '') },
        h('button', { type: 'button', class: 'pq-step-btn', onClick: () => { stepIndex = 0; renderAll(); } }, 'Source: ' + Q.sourceLabel(def.source)),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Change source', title: 'Change source (point this query at a different file, folder or sheet)', onClick: sourceDialog }, '⚙'))];
      def.steps.forEach((s, i) => {
        items.push(h('li', { class: 'pq-step' + (stepIndex === i + 1 ? ' active' : '') },
          h('button', { type: 'button', class: 'pq-step-btn', title: Q.describe(s), onClick: () => { stepIndex = i + 1; renderAll(); } }, Q.describe(s)),
          h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Delete step ' + Q.describe(s), title: 'Delete step', onClick: () => {
            def.steps.splice(i, 1);
            stepIndex = Math.min(stepIndex, def.steps.length);
            save();
          } }, '×')));
      });
      stepsList.replaceChildren(...items);
    }

    function renderPreview() {
      if (!def) { preview.replaceChildren(h('p', { class: 'muted pq-empty' }, 'Create a new query to get started: choose a file, a folder of files, or a sheet.')); msg.textContent = ''; return; }
      const res = current();
      msg.textContent = '';
      msg.className = 'pq-msg';
      if (res.error) {
        msg.textContent = 'Expression.Error: ' + res.error + (res.errorStep >= 0 ? ' (in step "' + Q.describe(def.steps[res.errorStep]) + '")' : '');
        msg.classList.add('err');
        if (!res.table) { preview.replaceChildren(); return; }
      }
      const t = res.table;
      if (selectedCol && !t.columns.includes(selectedCol)) selectedCol = null;
      const head = h('tr', {}, h('th', { class: 'pq-rownum' }, ''), t.columns.map((c) => h('th', {
        class: 'pq-col' + (c === selectedCol ? ' sel' : ''), tabindex: '0', title: 'Select column ' + c,
        onClick: () => { selectedCol = c; renderPreview(); renderRibbon(); },
        onKeydown: (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selectedCol = c; renderPreview(); renderRibbon(); } },
      }, h('span', { class: 'pq-type' }, TYPE_ICON[t.types[c]] || 'ABC123'), ' ', c)));
      const rows = t.rows.slice(0, 60).map((r, i) => h('tr', {}, h('td', { class: 'pq-rownum' }, String(i + 1)), t.columns.map((c) => {
        const v = r[c];
        const isE = F.isError(v);
        const text = v === null || v === undefined ? 'null' : isE ? 'Error' : t.types[c] === 'date' && typeof v === 'number' ? F.formatValue(v, 'date') : typeof v === 'number' ? F.formatGeneral(v) : String(v);
        return h('td', { class: (v === null || v === undefined ? 'null' : '') + (isE ? ' error' : '') + (c === selectedCol ? ' sel' : '') + (typeof v === 'number' ? ' num' : ''), title: isE ? v.hint : null }, text);
      })));
      const errNote = res.errorCells ? h('p', { class: 'pq-warn' }, res.errorCells + ' cell(s) show Error — a type conversion failed. Fix the data or use Remove Errors.') : null;
      preview.replaceChildren(
        h('p', { class: 'muted' }, t.columns.length + ' columns, ' + t.rows.length + ' rows' + (t.rows.length > 60 ? ' (showing first 60)' : '') + (stepIndex < def.steps.length ? ' — previewing step ' + stepIndex + ' of ' + def.steps.length : '')),
        errNote,
        h('div', { class: 'pq-table-wrap' }, h('table', { class: 'pq-table' }, h('thead', {}, head), h('tbody', {}, rows))));
    }

    function addStep(step) {
      if (!def) return;
      def.steps.splice(stepIndex, 0, step);
      stepIndex++;
      save();
    }

    function needCol() {
      if (!selectedCol) { toast('Click a column header in the preview first.', 'warn'); return false; }
      return true;
    }

    function columnsNow() { const r = current(); return r && r.table ? r.table.columns : []; }

    function rb(label, title, fn) { return h('button', { class: 'tb-btn', type: 'button', title, onClick: fn, disabled: !def }, label); }

    function renderRibbon() {
      const cols = columnsNow();
      ribbon.replaceChildren(
        h('div', { class: 'tb-group' },
          rb('Choose Columns', 'Keep only the columns you pick', () => columnsPicker('Choose Columns', cols, cols, (keep) => addStep({ type: 'keepColumns', columns: keep }))),
          rb('Remove Column', 'Remove the selected column', () => needCol() && addStep({ type: 'removeColumns', columns: [selectedCol] })),
          rb('Rename', 'Rename the selected column', () => { if (!needCol()) return; promptDialog('Rename column', 'New name', selectedCol, (to) => addStep({ type: 'renameColumn', from: selectedCol, to })); }),
          h('span', { class: 'tb-group-label' }, 'Columns')),
        h('div', { class: 'tb-group' },
          rb('Data Type', 'Change the data type of the selected column', () => { if (!needCol()) return; typeDialog(); }),
          rb('Replace Values', 'Replace values in the selected column', () => { if (!needCol()) return; replaceDialog(); }),
          rb('Format', 'Trim, clean or change the case of text', () => { if (!needCol()) return; formatDialog(); }),
          rb('Split Column', 'Split the selected column by a delimiter', () => { if (!needCol()) return; splitDialog(); }),
          rb('Fill Down', 'Fill blank cells with the value above', () => needCol() && addStep({ type: 'fillDown', column: selectedCol })),
          h('span', { class: 'tb-group-label' }, 'Transform')),
        h('div', { class: 'tb-group' },
          rb('Filter', 'Keep rows that match a condition', () => { if (!needCol()) return; filterDialog(); }),
          rb('Sort A→Z', 'Sort ascending by the selected column', () => needCol() && addStep({ type: 'sortRows', column: selectedCol, order: 'asc' })),
          rb('Sort Z→A', 'Sort descending by the selected column', () => needCol() && addStep({ type: 'sortRows', column: selectedCol, order: 'desc' })),
          rb('Remove Duplicates', 'Remove duplicate rows (based on the selected column, or all columns)', () => addStep({ type: 'removeDuplicates', columns: selectedCol ? [selectedCol] : [] })),
          rb('Remove Blank Rows', 'Remove rows where every value is blank', () => addStep({ type: 'removeBlankRows' })),
          rb('Remove Errors', 'Remove rows with an Error in the selected column', () => addStep({ type: 'removeErrors', column: selectedCol || undefined })),
          rb('Keep Top Rows', 'Keep only the first N rows', () => promptDialog('Keep Top Rows', 'Number of rows', '10', (n) => addStep({ type: 'keepTopRows', count: Number(n) }))),
          h('span', { class: 'tb-group-label' }, 'Rows')),
        h('div', { class: 'tb-group' },
          rb('Custom Column', 'Add a column calculated from others, e.g. [Qty] * [Price]', customDialog),
          rb('Group By', 'Summarize rows by one or more columns', groupDialog),
          rb('Unpivot Other Columns', 'Turn wide month columns into rows (keep the selected column)', () => { if (!needCol()) return; columnsPicker('Unpivot: keep these columns', cols, [selectedCol], (keep) => addStep({ type: 'unpivot', keep })); }),
          h('span', { class: 'tb-group-label' }, 'Shape')),
        h('div', { class: 'tb-group' },
          rb('Merge Queries', 'Join another query or file by a key column', mergeDialog),
          rb('Append Queries', 'Stack another query or file underneath', appendDialog),
          h('span', { class: 'tb-group-label' }, 'Combine')),
      );
    }

    // ---- small dialogs
    function promptDialog(title, label, value, done) {
      const input = h('input', { type: 'text', value });
      dialog(title, field(label, input), [{ label: 'Cancel' }, { label: 'OK', primary: true, onClick: () => { if (!input.value.trim()) return false; done(input.value.trim()); } }]);
    }

    function columnsPicker(title, cols, checked, done) {
      const list = h('div', { class: 'checklist' }, cols.map((c) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: c, checked: checked.includes(c) }), ' ' + c)));
      dialog(title, list, [{ label: 'Cancel' }, { label: 'OK', primary: true, onClick: () => {
        const pick = [...list.querySelectorAll('input:checked')].map((x) => x.value);
        if (!pick.length) { toast('Pick at least one column.', 'warn'); return false; }
        done(pick);
      } }]);
    }

    function typeDialog() {
      const s = select([['text', 'Text'], ['number', 'Decimal Number'], ['whole', 'Whole Number'], ['currency', 'Currency'], ['date', 'Date'], ['boolean', 'True/False']], 'number');
      dialog('Change Type: ' + selectedCol, field('Data type', s), [{ label: 'Cancel' }, { label: 'OK', primary: true, onClick: () => addStep({ type: 'changeType', columns: { [selectedCol]: s.value } }) }]);
    }

    function replaceDialog() {
      const find = h('input', { type: 'text' });
      const rep = h('input', { type: 'text' });
      const whole = h('input', { type: 'checkbox' });
      dialog('Replace Values in ' + selectedCol, h('div', {}, field('Value to find', find), field('Replace with', rep), h('label', { class: 'check' }, whole, ' Match entire cell contents')), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => addStep({ type: 'replaceValues', column: selectedCol, find: find.value, replace: rep.value, matchEntire: whole.checked }) },
      ]);
    }

    function formatDialog() {
      const s = select([['trim', 'Trim (remove extra spaces)'], ['clean', 'Clean (remove hidden characters)'], ['proper', 'Capitalize Each Word'], ['upper', 'UPPERCASE'], ['lower', 'lowercase']], 'trim');
      dialog('Format ' + selectedCol, field('Transform', s), [{ label: 'Cancel' }, { label: 'OK', primary: true, onClick: () => addStep({ type: 'transform', column: selectedCol, op: s.value }) }]);
    }

    function splitDialog() {
      const s = select([[',', 'Comma'], [' ', 'Space'], ['-', 'Dash'], ['|', 'Pipe'], [';', 'Semicolon'], ['/', 'Slash']], ',');
      dialog('Split Column by Delimiter', field('Delimiter', s), [{ label: 'Cancel' }, { label: 'OK', primary: true, onClick: () => addStep({ type: 'splitColumn', column: selectedCol, delimiter: s.value }) }]);
    }

    function filterDialog() {
      const op = select([['=', 'equals'], ['<>', 'does not equal'], ['>', 'is greater than'], ['>=', 'is greater than or equal to'], ['<', 'is less than'], ['<=', 'is less than or equal to'], ['contains', 'contains'], ['beginsWith', 'begins with'], ['notBlank', 'is not blank'], ['blank', 'is blank']], '=');
      const val = h('input', { type: 'text' });
      dialog('Filter Rows: ' + selectedCol, h('div', {}, field('Keep rows where ' + selectedCol, op), field('Value', val)), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          const p = F.parseNumberLike(val.value);
          addStep({ type: 'filterRows', column: selectedCol, op: op.value, value: p && p.format !== 'date' ? p.value : val.value });
        } },
      ]);
    }

    function customDialog() {
      const name = h('input', { type: 'text', value: 'Custom' });
      const formula = h('input', { type: 'text', placeholder: '[Qty] * [Price]' });
      const cols = columnsNow();
      const chips = h('div', { class: 'chips' }, cols.map((c) => h('button', { type: 'button', class: 'chip', onClick: () => { formula.value += '[' + c + ']'; formula.focus(); } }, c)));
      dialog('Add Custom Column', h('div', {}, field('New column name', name), field('Formula', formula, 'Click a column to insert it. Excel functions work too, e.g. ROUND([Price]*1.08, 2) or UPPER([Name]).'), chips), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          try { Q.compileCustom(formula.value, cols); } catch (e) { toast('Formula problem: ' + e.message, 'warn'); return false; }
          addStep({ type: 'addColumn', name: name.value.trim(), formula: formula.value });
        } },
      ]);
    }

    function groupDialog() {
      const cols = columnsNow();
      if (!cols.length) return;
      const by = select(cols, selectedCol || cols[0]);
      const op = select([['sum', 'Sum'], ['count', 'Count Rows'], ['average', 'Average'], ['min', 'Min'], ['max', 'Max'], ['countDistinct', 'Count Distinct']], 'sum');
      const col = select(cols, cols[cols.length - 1]);
      const name = h('input', { type: 'text', value: 'Total' });
      dialog('Group By', h('div', {}, field('Group by', by), field('New column name', name), field('Operation', op), field('Column', col)), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => addStep({ type: 'groupBy', by: [by.value], aggregations: [{ name: name.value.trim() || 'Total', op: op.value, column: col.value }] }) },
      ]);
    }

    function otherSources() {
      const list = wb.queries.filter((q) => !def || q.id !== def.id).map((q) => ['query:' + q.name, 'Query: ' + q.name]);
      Object.keys(wb.sources || {}).forEach((f) => list.push(['file:' + f, 'File: ' + f]));
      wb.sheets.forEach((s) => list.push(['sheet:' + s.name, 'Sheet: ' + s.name]));
      return list;
    }
    function parseSource(v) {
      const [kind, ...rest] = v.split(':');
      const name = rest.join(':');
      if (kind === 'query') return { kind: 'query', name };
      if (kind === 'file') return { kind: 'file', file: name };
      return { kind: 'sheet', sheet: name };
    }

    function mergeDialog() {
      const cols = columnsNow();
      const srcSel = select(otherSources(), (otherSources()[0] || [''])[0]);
      const leftKey = select(cols, selectedCol || cols[0]);
      const rightKey = h('select', {});
      const expand = h('div', { class: 'checklist' });
      const kind = select([['leftOuter', 'Left Outer (all from first, matching from second)'], ['inner', 'Inner (only matching rows)'], ['fullOuter', 'Full Outer (all rows from both)'], ['rightOuter', 'Right Outer (all from second, matching from first)'], ['leftAnti', 'Left Anti (rows only in first)'], ['rightAnti', 'Right Anti (rows only in second)']], 'leftOuter');
      const fillRight = () => {
        const r = Q.loadSource(parseSource(srcSel.value), wb);
        const rc = r.table ? r.table.columns : [];
        rightKey.replaceChildren(...rc.map((c) => h('option', { value: c, selected: c.toLowerCase() === String(leftKey.value).toLowerCase() }, c)));
        expand.replaceChildren(...rc.map((c) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: c, checked: true }), ' ' + c)));
      };
      srcSel.addEventListener('change', fillRight);
      fillRight();
      dialog('Merge Queries', h('div', {},
        h('p', { class: 'muted' }, 'Match rows where the key in this query equals the key in the other table. This is a join, the same idea as XLOOKUP.'),
        field('Key column in this query', leftKey), field('Merge with', srcSel), field('Matching key column', rightKey), field('Join kind', kind),
        h('p', { class: 'field-label' }, 'Columns to bring in'), expand), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          const pick = [...expand.querySelectorAll('input:checked')].map((x) => x.value).filter((c) => c !== rightKey.value);
          addStep({ type: 'merge', with: parseSource(srcSel.value), leftKey: leftKey.value, rightKey: rightKey.value, kind: kind.value, columns: pick });
        } },
      ], { wide: true });
    }

    function appendDialog() {
      const list = h('div', { class: 'checklist' }, otherSources().map(([v, l]) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: v }), ' ' + l)));
      dialog('Append Queries', h('div', {}, h('p', { class: 'muted' }, 'Stack rows from these sources under this query. Columns are matched by name.'), list), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          const pick = [...list.querySelectorAll('input:checked')].map((x) => parseSource(x.value));
          if (!pick.length) { toast('Pick at least one source.', 'warn'); return false; }
          addStep({ type: 'append', with: pick });
        } },
      ]);
    }

    // Point the query at a new file (e.g. next month's export). Every applied
    // step re-runs on the new data: the point of Power Query.
    function sourceDialog() {
      const files = Object.keys(wb.sources || {});
      const src = def.source || {};
      let body, read;
      if (src.kind === 'file') {
        const sel = select(files, src.file);
        body = field('File', sel, 'Every applied step will run again on the new file.');
        read = () => ({ ...src, file: sel.value });
      } else if (src.kind === 'folder') {
        const list = h('div', { class: 'checklist' }, files.map((f) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: f, checked: (src.files || []).includes(f) }), ' ' + f)));
        body = h('div', {}, h('p', { class: 'field-label' }, 'Files in the folder'), list);
        read = () => { const pick = [...list.querySelectorAll('input:checked')].map((x) => x.value); return pick.length ? { ...src, files: pick } : null; };
      } else if (src.kind === 'sheet') {
        const sel = select(wb.sheets.map((x) => x.name), src.sheet);
        body = field('Sheet', sel);
        read = () => ({ ...src, sheet: sel.value });
      } else return;
      dialog('Change Source', body, [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          const next = read();
          if (!next) { toast('Pick at least one file.', 'warn'); return false; }
          def.source = next;
          save();
        } },
      ]);
    }

    function newQueryDialog() {
      const files = Object.keys(wb.sources || {});
      const kind = select([['file', 'From Text/CSV file'], ['folder', 'From Folder (combine several files)'], ['sheet', 'From a sheet in this workbook']], files.length ? 'file' : 'sheet');
      const fileSel = select(files.length ? files : [['', 'No files provided']], files[0]);
      const folder = h('div', { class: 'checklist' }, files.map((f) => h('label', { class: 'check' }, h('input', { type: 'checkbox', value: f, checked: true }), ' ' + f)));
      const sheetSel = select(wb.sheets.map((s) => s.name), wb.sheets[0].name);
      const autoType = h('input', { type: 'checkbox', checked: true });
      const name = h('input', { type: 'text', value: '' });
      const fFile = field('File', fileSel), fFolder = h('div', {}, h('p', { class: 'field-label' }, 'Files in the folder'), folder), fSheet = field('Sheet', sheetSel);
      const upd = () => { fFile.style.display = kind.value === 'file' ? '' : 'none'; fFolder.style.display = kind.value === 'folder' ? '' : 'none'; fSheet.style.display = kind.value === 'sheet' ? '' : 'none'; };
      kind.addEventListener('change', upd); upd();
      dialog('Get Data', h('div', {}, field('Source', kind), fFile, fFolder, fSheet, field('Query name', name), h('label', { class: 'check' }, autoType, ' Detect data types (adds a "Changed Type" step)')), [
        { label: 'Cancel' },
        { label: 'OK', primary: true, onClick: () => {
          let source;
          if (kind.value === 'file') { if (!fileSel.value) return false; source = { kind: 'file', file: fileSel.value }; }
          else if (kind.value === 'folder') {
            const pick = [...folder.querySelectorAll('input:checked')].map((x) => x.value);
            if (!pick.length) { toast('Pick at least one file.', 'warn'); return false; }
            source = { kind: 'folder', files: pick };
          } else source = { kind: 'sheet', sheet: sheetSel.value };
          let qname = (name.value.trim() || (kind.value === 'file' ? fileSel.value.replace(/\.\w+$/, '') : kind.value === 'folder' ? 'Combined' : sheetSel.value + ' Query')).replace(/[^\w ]/g, '').slice(0, 28) || 'Query';
          let n = 2; const base = qname;
          while (wb.queries.some((q) => q.name.toLowerCase() === qname.toLowerCase()) || (wb.getSheet(qname) && !wb.queries.some((q) => q.name === qname))) qname = base + ' ' + n++;
          def = { id: 'q' + Math.random().toString(36).slice(2, 8), name: qname, source, steps: [] };
          if (autoType.checked) {
            const r = Q.loadSource(source, wb);
            if (r.table) {
              const types = detectTypes(r.table);
              if (Object.keys(types).length) def.steps.push({ type: 'changeType', columns: types });
            }
          }
          stepIndex = def.steps.length;
          nameIn.value = def.name;
          selectedCol = null;
          save();
        } },
      ]);
    }

    function detectTypes(t) {
      const out = {};
      t.columns.forEach((c) => {
        const vals = t.rows.map((r) => r[c]).filter((v) => v !== null && v !== undefined && v !== '').slice(0, 50);
        if (!vals.length) return;
        if (vals.every((v) => typeof v === 'number')) return;
        const parsed = vals.map((v) => F.parseNumberLike(String(v)));
        if (vals.some((v) => /^0\d/.test(String(v).trim()))) { out[c] = 'text'; return; } // keep leading zeros (IDs, zip codes)
        if (parsed.every((p) => p && p.format === 'date')) out[c] = 'date';
        else if (parsed.every((p) => p && p.format !== 'date')) out[c] = parsed.some((p) => p.format === 'currency') ? 'currency' : parsed.every((p) => Number.isInteger(p.value)) ? 'whole' : 'number';
        else out[c] = 'text';
      });
      return out;
    }

    if (opts.queryId) {
      const q = wb.queries.find((x) => x.id === opts.queryId);
      if (q) { def = JSON.parse(JSON.stringify(q)); stepIndex = def.steps.length; nameIn.value = def.name; }
    } else if (wb.queries.length) {
      def = JSON.parse(JSON.stringify(wb.queries[0]));
      stepIndex = def.steps.length;
      nameIn.value = def.name;
    }
    renderAll();
    return d;
  }

  root.DIM = root.DIM || {};
  root.DIM.QueryEditor = { open };
})(typeof globalThis !== 'undefined' ? globalThis : this);
