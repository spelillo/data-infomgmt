/*
 * grader.js — automatic grading for quizzes, exams, assignments and practice.
 *
 * An assessment has an optional workbook (the "template" students work in) and
 * a list of questions. Workbook questions are checked against the student's
 * workbook; the others (multiple choice, numeric, short text, matching) are
 * checked against the answers typed on the page.
 *
 * Expected values for workbook questions are not hard-coded: the grader builds
 * an answer key by applying every question's `solution` to a fresh copy of the
 * template and recalculating. That keeps answer keys correct by construction
 * and lets the test suite prove every quiz can be solved for 100%.
 *
 * Question types (see docs/AUTHORING.md for the full reference):
 *   cell        one cell: value (+ optional formula / function / reference rules)
 *   range       a block filled from one formula, e.g. D2:D11 from =B2*C2
 *   cellType    a cell or range must hold numbers / text / dates (and a value)
 *   filled      at least N cells in a range must contain something (open-ended work)
 *   format      cells must use a number format (currency, percent, date...)
 *   table       an Excel Table must exist over a range
 *   validation  a data-validation rule must exist on a range
 *   condFormat  a conditional-formatting rule must exist on a range
 *   sorted      a block must be sorted by a column, with its rows intact
 *   chart       a chart of a given type must exist
 *   pivot       a PivotTable with given fields must exist
 *   query       a Power Query with given step types must exist (and load)
 *   choice      multiple choice (single or multi-select)
 *   number      a typed numeric answer
 *   text        a typed short answer, matched against accepted answers
 *   match       match each term to its definition
 */
(function (root) {
  'use strict';

  const isNode = typeof module === 'object' && module.exports;
  const F = isNode ? require('./formula.js') : root.DIM.Formula;
  const W = isNode ? require('./workbook.js') : root.DIM.Workbook;

  const WORKBOOK_TYPES = new Set(['cell', 'range', 'cellType', 'filled', 'format', 'table', 'validation', 'condFormat', 'sorted', 'chart', 'pivot', 'query']);

  function isWorkbookQuestion(q) { return WORKBOOK_TYPES.has(q.type); }

  function sheetOf(q, assessment) {
    return q.sheet || (assessment.workbook && assessment.workbook.sheets && assessment.workbook.sheets[0] && assessment.workbook.sheets[0].name) || 'Sheet1';
  }

  // ---------------------------------------------------------------------------
  // Answer key
  // ---------------------------------------------------------------------------

  // Apply one question's solution to a workbook (used for keys and for tests).
  function applySolution(wb, q, assessment) {
    const sheet = sheetOf(q, assessment);
    let sol = q.solution;
    // solution: null means the assessment's model walkthrough already produces it
    if (sol === undefined) {
      // Fall back to what the question already states it expects.
      if (q.type === 'sorted') sol = true;
      else if (q.type === 'format') sol = q.format;
      else if (q.type === 'table') sol = q.name || true;
      else if (q.type === 'filled') sol = 'Sample answer';
      else if ((q.type === 'cell' || q.type === 'range' || q.type === 'cellType') && q.expect !== undefined) sol = expectAsInputs(q);
    }
    if (sol === undefined || sol === null) return;
    if (q.type === 'cell' || q.type === 'cellType') {
      if (q.cell) wb.setInput(sheet, q.cell, sol, { force: true });
      else if (q.range && Array.isArray(sol)) {
        const r = W.normRange(q.range);
        const cols = r.c2 - r.c1 + 1;
        sol.flat().forEach((v, i) => wb.setInput(sheet, { row: r.r1 + Math.floor(i / cols), col: r.c1 + (i % cols) }, v, { force: true }));
      } else if (q.range) {
        const r = W.normRange(q.range);
        wb.setInput(sheet, { row: r.r1, col: r.c1 }, sol, { force: true });
        if (r.r2 > r.r1 || r.c2 > r.c1) wb.fill(sheet, q.range, r.r2 > r.r1 ? 'down' : 'right');
      }
      if (q.type === 'cellType' && q.valueType === 'date') wb.setFormat(sheet, q.cell || q.range, 'date');
    } else if (q.type === 'filled') {
      if (!wb.getSheet(sheet)) wb.addSheet(sheet, { silent: true });
      const r = W.normRange(q.range);
      let n = 0;
      for (let row = r.r1; row <= r.r2 && n < (q.min || 1); row++) for (let col = r.c1; col <= r.c2 && n < (q.min || 1); col++, n++) {
        wb.setInput(sheet, { row, col }, Array.isArray(sol) ? sol[n % sol.length] : sol, { force: true });
      }
    } else if (q.type === 'range') {
      const r = W.normRange(q.range);
      if (Array.isArray(sol)) {
        sol.flat().forEach((v, i) => {
          const cols = r.c2 - r.c1 + 1;
          wb.setInput(sheet, { row: r.r1 + Math.floor(i / cols), col: r.c1 + (i % cols) }, v, { force: true });
        });
      } else {
        const s = wb.getSheet(sheet);
        for (let i = r.r1; i <= r.r2; i++) for (let j = r.c1; j <= r.c2; j++) {
          const key = F.toAddress(i, j);
          if (s.cells[key]) s.cells[key].locked = false;
        }
        wb.setInput(sheet, { row: r.r1, col: r.c1 }, sol, { force: true });
        if (r.r2 > r.r1) wb.fill(sheet, F.rangeToAddress({ r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c1 }), 'down');
        if (r.c2 > r.c1) wb.fill(sheet, F.rangeToAddress(r), 'right');
      }
    } else if (q.type === 'format') {
      wb.setFormat(sheet, q.range || q.cell, sol);
    } else if (q.type === 'table') {
      wb.addTable(sheet, q.range, typeof sol === 'string' ? sol : undefined);
    } else if (q.type === 'validation') {
      wb.addValidation(sheet, q.range, sol);
    } else if (q.type === 'condFormat') {
      wb.addCondFormat(sheet, q.range, sol);
    } else if (q.type === 'sorted') {
      const r = W.normRange(q.range);
      wb.sortRange(sheet, q.range, F.colToIndex(q.column) - r.c1, q.order !== 'desc', true);
    } else if (q.type === 'chart') {
      wb.addChart({ sheet, ...sol });
    } else if (q.type === 'pivot') {
      wb.addPivot({ sourceSheet: sheet, ...sol });
    } else if (q.type === 'query') {
      const def = { id: 'q-' + (sol.name || 'Query'), ...sol };
      wb.saveQuery(def);
      wb.loadQuery(def.id);
    }
  }

  // Turn a question's `expect` into inputs a student could type: text that
  // Excel would read as a number or date gets a leading apostrophe.
  function expectAsInputs(q) {
    const conv = W.inputForValue;
    return Array.isArray(q.expect) ? q.expect.map((x) => (Array.isArray(x) ? x.map(conv) : conv(x))) : conv(q.expect);
  }

  function buildTemplate(assessment) {
    return W.Workbook.fromSpec(assessment.workbook || { sheets: [{ name: 'Sheet1' }] });
  }

  // The answer key: the template, then the author's optional model walkthrough
  // (steps no single question's solution can express, such as importing a file
  // or removing duplicates), then every question's solution.
  function buildKey(assessment) {
    const wb = buildTemplate(assessment);
    if (typeof assessment.model === 'function') assessment.model(wb, { W, F });
    const qs = (assessment.questions || []).filter(isWorkbookQuestion);
    // apply in order; structural ones (queries, pivots, sorts) first so cell answers see them
    const order = ['query', 'table', 'pivot', 'sorted', 'format', 'validation', 'condFormat', 'chart', 'cellType', 'cell', 'range'];
    qs.slice().sort((a, b) => order.indexOf(a.type) - order.indexOf(b.type)).forEach((q) => applySolution(wb, q, assessment));
    wb.recalc();
    return wb;
  }

  // ---------------------------------------------------------------------------
  // Comparison helpers
  // ---------------------------------------------------------------------------

  function show(v, fmt) {
    if (v === null || v === undefined || v === '') return '(blank)';
    if (F.isError(v)) return v.code;
    if (typeof v === 'string') return '"' + v + '"';
    return F.formatValue(v, fmt);
  }

  function valuesMatch(actual, expected, q) {
    if (F.isError(expected)) return F.isError(actual) && actual.code === expected.code;
    if (F.isError(actual)) return false;
    const blank = (v) => v === null || v === undefined || v === '';
    if (blank(expected)) return blank(actual);
    if (typeof expected === 'number') {
      if (typeof actual !== 'number') return false;
      const tol = q.tolerance !== undefined ? q.tolerance : Math.max(1e-9, Math.abs(expected) * 1e-9);
      return Math.abs(actual - expected) <= tol;
    }
    if (typeof expected === 'boolean') return actual === expected;
    if (typeof actual !== 'string') return false;
    const norm = (s) => (q.exactText ? s : s.trim().replace(/\s+/g, ' ').toLowerCase());
    if (q.caseSensitive) return actual.trim() === expected.trim();
    return norm(actual) === norm(expected);
  }

  // Rules about *how* the answer was produced.
  function methodIssues(wb, sheet, addr, q) {
    const issues = [];
    const input = wb.getInput(sheet, addr);
    const isFormula = input.startsWith('=') && input.length > 1;
    if ((q.requireFormula || (q.functions && q.functions.length) || q.mustReference) && !isFormula) {
      issues.push({ penalty: q.formulaPenalty !== undefined ? q.formulaPenalty : 0.5, msg: 'the value is typed in, but this task asks for a formula' });
      return issues;
    }
    if (q.functions && q.functions.length && isFormula) {
      let used;
      try { used = F.functionsUsed(F.parse(input)); } catch (e) { used = new Set(); }
      const missing = q.functions.filter((f) => ![].concat(f).some((alt) => used.has(alt.toUpperCase())));
      if (missing.length) issues.push({ penalty: 0.25, msg: 'use ' + missing.map((f) => [].concat(f).join(' or ')).join(' and ') + ' in your formula' });
    }
    if (q.mustReference && isFormula) {
      const body = input.toUpperCase().replace(/\s/g, '');
      const missing = [].concat(q.mustReference).filter((ref) => !body.includes(String(ref).toUpperCase()));
      if (missing.length) issues.push({ penalty: 0.25, msg: 'your formula should reference ' + missing.join(', ') });
    }
    if (q.forbidFunctions && isFormula) {
      let used;
      try { used = F.functionsUsed(F.parse(input)); } catch (e) { used = new Set(); }
      const bad = q.forbidFunctions.filter((f) => used.has(f));
      if (bad.length) issues.push({ penalty: 0.25, msg: 'solve this without ' + bad.join(', ') });
    }
    return issues;
  }

  // ---------------------------------------------------------------------------
  // Per-question graders
  // ---------------------------------------------------------------------------

  function missingSheet(q, sheet) {
    return result(q, 0, 'There is no sheet named "' + sheet + '" yet. Check the sheet name (double-click a tab to rename it).');
  }

  function result(q, earnedFrac, feedback, extra) {
    const points = q.points !== undefined ? q.points : 1;
    const earned = Math.round(points * Math.max(0, Math.min(1, earnedFrac)) * 100) / 100;
    return {
      id: q.id, type: q.type, points, earned,
      status: earned >= points ? 'correct' : earned > 0 ? 'partial' : 'incorrect',
      feedback, ...(extra || {}),
    };
  }

  function gradeCell(q, wb, key, assessment, opts) {
    const sheet = sheetOf(q, assessment);
    if (!wb.getSheet(sheet)) return missingSheet(q, sheet);
    const addr = q.cell;
    const actual = wb.getValue(sheet, addr);
    const expected = q.expect !== undefined ? q.expect : key.getValue(sheet, addr);
    const fmt = key.getFormat(sheet, addr) || wb.getFormat(sheet, addr);
    if (!valuesMatch(actual, expected, q)) {
      const blank = actual === null || actual === undefined || actual === '';
      const fb = blank
        ? addr + ' is empty.'
        : addr + ' shows ' + show(actual, fmt) + (opts.reveal ? ', expected ' + show(expected, fmt) + '.' : ', which is not correct yet.');
      return result(q, 0, fb + (q.hint && !blank ? ' Hint: ' + q.hint : ''), { cell: addr, expected: opts.reveal ? show(expected, fmt) : undefined });
    }
    const issues = methodIssues(wb, sheet, addr, q);
    if (issues.length) {
      const frac = 1 - issues.reduce((a, b) => a + b.penalty, 0);
      return result(q, frac, addr + ' has the right value, but ' + issues.map((i) => i.msg).join('; ') + '.', { cell: addr });
    }
    return result(q, 1, addr + ' is correct.', { cell: addr });
  }

  function gradeRange(q, wb, key, assessment, opts) {
    const sheet = sheetOf(q, assessment);
    if (!wb.getSheet(sheet)) return missingSheet(q, sheet);
    const r = W.normRange(q.range);
    let right = 0, total = 0, methodPenalty = 0;
    const wrong = [];
    const methodMsgs = new Set();
    const expected = Array.isArray(q.expect) ? q.expect.flat() : null;
    let i = 0;
    for (let row = r.r1; row <= r.r2; row++) {
      for (let col = r.c1; col <= r.c2; col++, i++) {
        total++;
        const addr = F.toAddress(row, col);
        const exp = expected ? expected[i] : key.getValue(sheet, addr);
        const act = wb.getValue(sheet, addr);
        if (valuesMatch(act, exp, q)) {
          right++;
          const issues = methodIssues(wb, sheet, addr, q);
          if (issues.length) { methodPenalty += issues.reduce((a, b) => a + b.penalty, 0); issues.forEach((x) => methodMsgs.add(x.msg)); }
        } else wrong.push(addr + (opts.reveal ? ' (got ' + show(act) + ', expected ' + show(exp) + ')' : ''));
      }
    }
    const frac = total ? (right - methodPenalty) / total : 0;
    let fb;
    if (!wrong.length && !methodMsgs.size) fb = 'All ' + total + ' cells in ' + q.range + ' are correct.';
    else {
      fb = right + ' of ' + total + ' cells in ' + q.range + ' are correct.';
      if (wrong.length) fb += ' Check ' + wrong.slice(0, 4).join(', ') + (wrong.length > 4 ? ' and ' + (wrong.length - 4) + ' more' : '') + '.';
      if (methodMsgs.size) fb += ' Also: ' + [...methodMsgs].join('; ') + '.';
      if (q.hint && wrong.length) fb += ' Hint: ' + q.hint;
    }
    return result(q, frac, fb, { range: q.range });
  }

  function gradeCellType(q, wb, key, assessment) {
    const sheet = sheetOf(q, assessment);
    if (!wb.getSheet(sheet)) return missingSheet(q, sheet);
    const r = W.normRange(q.cell || q.range);
    const want = q.valueType; // 'number' | 'text' | 'date'
    let ok = 0, total = 0;
    const bad = [];
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
      total++;
      const addr = F.toAddress(row, col);
      const v = wb.getValue(sheet, addr);
      const fmt = wb.getFormat(sheet, addr);
      let typeOk;
      if (want === 'number') typeOk = typeof v === 'number';
      else if (want === 'text') typeOk = typeof v === 'string';
      else if (want === 'date') typeOk = typeof v === 'number' && (fmt === 'date' || fmt === 'longdate');
      else typeOk = true;
      const exp = q.checkValue === false ? undefined
        : q.expect !== undefined ? (Array.isArray(q.expect) ? [].concat(...q.expect.map((x) => (Array.isArray(x) ? x : [x])))[total - 1] : q.expect)
        : key.getValue(sheet, addr);
      const valueOk = exp === undefined || exp === null ? true : valuesMatch(v, exp, q);
      if (typeOk && valueOk) ok++; else bad.push(addr);
    }
    const label = { number: 'numbers', text: 'text', date: 'real dates (date format)' }[want] || want;
    const fb = bad.length
      ? bad.slice(0, 5).join(', ') + (bad.length > 5 ? ' and ' + (bad.length - 5) + ' more' : '') + (bad.length === 1 ? ' is' : ' are') + ' not stored as ' + label + ' with the right value yet.' + (q.hint ? ' Hint: ' + q.hint : '')
      : 'Every cell is stored as ' + label + '.';
    return result(q, total ? ok / total : 0, fb);
  }

  function gradeFilled(q, wb, key, assessment) {
    const sheet = sheetOf(q, assessment);
    const r = W.normRange(q.range);
    const s = wb.getSheet(sheet);
    if (!s) return result(q, 0, 'Sheet "' + sheet + '" was not found. Check the sheet name.');
    let filled = 0;
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
      const v = wb.getValue(sheet, { row, col });
      if (v !== null && v !== '' && !(F.isError(v) && v.code === '#REF!')) filled++;
    }
    const need = q.min || 1;
    return result(q, Math.min(1, filled / need), filled >= need ? 'Done: ' + filled + ' cell(s) filled in ' + q.range + '.' : 'Only ' + filled + ' of the ' + need + ' required cells in ' + q.range + ' are filled.');
  }

  function gradeFormat(q, wb, key, assessment) {
    const sheet = sheetOf(q, assessment);
    const r = W.normRange(q.range || q.cell);
    const want = [].concat(q.solution || q.format);
    let ok = 0, total = 0;
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
      total++;
      if (want.includes(wb.getFormat(sheet, { row, col }))) ok++;
    }
    return result(q, ok / total, ok === total ? (q.range || q.cell) + ' uses the ' + want[0] + ' format.' : ok + ' of ' + total + ' cells use the ' + want[0] + ' format.');
  }

  function overlaps(a, b) { return !(a.r2 < b.r1 || a.r1 > b.r2 || a.c2 < b.c1 || a.c1 > b.c2); }
  function covers(outer, inner) { return outer.r1 <= inner.r1 && outer.c1 <= inner.c1 && outer.r2 >= inner.r2 && outer.c2 >= inner.c2; }

  function gradeTable(q, wb, key, assessment) {
    const sheet = sheetOf(q, assessment);
    const want = W.normRange(q.range);
    const t = wb.tables.find((x) => x.sheet.toLowerCase() === sheet.toLowerCase() && covers(W.normRange(x.ref), want));
    if (!t) {
      const near = wb.tables.find((x) => x.sheet.toLowerCase() === sheet.toLowerCase() && overlaps(W.normRange(x.ref), want));
      return result(q, near ? 0.5 : 0, near ? 'Table ' + near.name + ' only covers ' + near.ref + '; it should cover ' + q.range + '.' : 'No Excel Table covers ' + q.range + '. Select the data and use Format as Table.');
    }
    if (q.name && t.name.toLowerCase() !== String(q.name).toLowerCase()) return result(q, 0.75, 'The table exists but is named ' + t.name + ' instead of ' + q.name + '.');
    return result(q, 1, 'Table ' + t.name + ' covers ' + t.ref + '.');
  }

  function gradeValidation(q, wb, key, assessment) {
    const sheet = wb.getSheet(sheetOf(q, assessment));
    const want = W.normRange(q.range);
    const rule = (sheet.validations || []).find((v) => covers(W.normRange(v.range), want));
    if (!rule) return result(q, 0, 'No data validation rule covers ' + q.range + '.');
    const sol = q.solution || {};
    if (sol.type && rule.type !== sol.type) return result(q, 0.5, 'There is a rule on ' + q.range + ', but it should be a ' + sol.type + ' rule.');
    if (sol.type === 'list' && sol.values) {
      const have = (rule.values || []).map((x) => String(x).toLowerCase()).sort();
      const need = sol.values.map((x) => String(x).toLowerCase()).sort();
      if (have.join('|') !== need.join('|')) return result(q, 0.5, 'The list should contain exactly: ' + sol.values.join(', ') + '.');
    }
    if ((sol.type === 'whole' || sol.type === 'decimal') && (rule.min !== sol.min || rule.max !== sol.max)) {
      return result(q, 0.5, 'The rule should allow ' + sol.min + ' to ' + sol.max + '.');
    }
    return result(q, 1, 'Data validation is set on ' + q.range + '.');
  }

  function gradeCondFormat(q, wb, key, assessment) {
    const sheet = wb.getSheet(sheetOf(q, assessment));
    const want = W.normRange(q.range);
    const sol = q.solution || {};
    const rule = (sheet.condFormats || []).find((cf) => overlaps(W.normRange(cf.range), want) && (!sol.op || cf.op === sol.op));
    if (!rule) return result(q, 0, 'No matching conditional formatting rule on ' + q.range + '.');
    if (!covers(W.normRange(rule.range), want)) return result(q, 0.5, 'The rule should cover all of ' + q.range + '.');
    if (sol.value !== undefined && String(rule.value) !== String(sol.value)) return result(q, 0.5, 'The rule should use the value ' + sol.value + '.');
    return result(q, 1, 'Conditional formatting is applied to ' + q.range + '.');
  }

  function gradeSorted(q, wb, key, assessment) {
    const sheet = sheetOf(q, assessment);
    const r = W.normRange(q.range);
    const colIdx = F.colToIndex(q.column) - r.c1;
    const rows = wb.getRangeValues(sheet, { ...r, r1: r.r1 + 1 });
    const keyRows = key.getRangeValues(sheet, { ...r, r1: r.r1 + 1 });
    const sig = (rs) => rs.map((row) => row.map((v) => (F.isError(v) ? v.code : String(v))).join('\u0001')).sort().join('\n');
    if (sig(rows) !== sig(keyRows)) return result(q, 0, 'The rows in ' + q.range + ' were changed or split apart. Sort whole rows so each record stays together (Undo and try again).');
    const desc = q.order === 'desc';
    for (let i = 1; i < rows.length; i++) {
      const c = F.compare(rows[i - 1][colIdx], rows[i][colIdx]);
      if ((desc && c < 0) || (!desc && c > 0)) return result(q, 0, 'The data is not sorted ' + (desc ? 'largest to smallest (Z→A)' : 'smallest to largest (A→Z)') + ' by column ' + q.column + ' yet.');
    }
    return result(q, 1, 'Sorted correctly by column ' + q.column + '.');
  }

  function gradeChart(q, wb, key, assessment) {
    const sheet = sheetOf(q, assessment);
    const sol = q.solution || {};
    const types = [].concat(q.chartType || sol.type);
    const charts = wb.charts.filter((c) => !q.sheet || c.sheet.toLowerCase() === sheet.toLowerCase());
    if (!charts.length) return result(q, 0, 'No chart found. Select the data and use Insert Chart.');
    const right = charts.find((c) => types.includes(c.type));
    if (!right) return result(q, 0.25, 'You made a ' + charts[0].type + ' chart, but a ' + types.join(' or ') + ' chart fits this question better.');
    let frac = 1;
    const notes = [];
    if (sol.range && !q.anyRange) {
      const want = W.normRange(sol.range);
      if (!covers(W.normRange(right.range), want) && !covers(want, W.normRange(right.range))) { frac -= 0.25; notes.push('the chart should use ' + sol.range); }
    }
    if (q.requireTitle && !(right.title || '').trim()) { frac -= 0.25; notes.push('give the chart a title'); }
    if (q.axisFromZero && right.axisMin !== null && right.axisMin !== undefined && right.axisMin !== 0) { frac -= 0.25; notes.push('start the value axis at 0 so bars are not exaggerated'); }
    return result(q, frac, notes.length ? 'Chart found, but ' + notes.join('; ') + '.' : 'A ' + right.type + ' chart is in place.');
  }

  function gradePivot(q, wb, key, assessment) {
    const sol = q.solution || {};
    const eq = (a, b) => !a || !b ? !a && !b : String(a).toLowerCase() === String(b).toLowerCase();
    const p = wb.pivots.find((x) => eq(x.rows, sol.rows) && eq(x.values, sol.values));
    if (!p) {
      if (!wb.pivots.length) return result(q, 0, 'No PivotTable found. Use Insert PivotTable.');
      return result(q, 0.25, 'Your PivotTable should have ' + sol.rows + ' in Rows and ' + sol.values + ' in Values.');
    }
    let frac = 1;
    const notes = [];
    if ((p.agg || 'sum') !== (sol.agg || 'sum')) { frac -= 0.5; notes.push('summarize Values by ' + (sol.agg || 'sum')); }
    if (sol.columns !== undefined && !eq(p.columns, sol.columns)) { frac -= 0.25; notes.push(sol.columns ? 'put ' + sol.columns + ' in Columns' : 'leave Columns empty'); }
    if (sol.rowsGroup && p.rowsGroup !== sol.rowsGroup) { frac -= 0.25; notes.push('group the row dates by ' + sol.rowsGroup); }
    return result(q, frac, notes.length ? 'PivotTable found, but ' + notes.join(' and ') + '.' : 'PivotTable is set up correctly.');
  }

  function gradeQuery(q, wb, key) {
    const sol = q.solution || {};
    const base = (f) => String(f || '').split(/[\\/]/).pop().toLowerCase();
    const sameImported = (x) => x.imported && x.source && sol.source && x.source.kind === sol.source.kind
      && (x.source.kind === 'folder' || (x.source.kind === 'file' && base(x.source.file) === base(sol.source.file))
        || (x.source.kind === 'sheet' && String(x.source.sheet).toLowerCase() === String(sol.source.sheet).toLowerCase())
        || (x.source.kind === 'query' && String(x.source.name).toLowerCase() === String(sol.source.name).toLowerCase()));
    const sameSource = (x) => sameImported(x) || JSON.stringify(x.source) === JSON.stringify(sol.source)
      || (x.source && sol.source && x.source.kind === 'folder' && sol.source.kind === 'folder' && [...x.source.files].sort().join('|') === [...sol.source.files].sort().join('|'));
    const cand = wb.queries.filter((x) => !sol.source || sameSource(x) || (!q.requireSource && x.source && x.source.kind === sol.source.kind));
    if (!cand.length && q.requireSource && wb.queries.length) return result(q, 0, 'Your query uses a different source. Point it at ' + (root.DIM && root.DIM.Query ? root.DIM.Query.sourceLabel(sol.source) : JSON.stringify(sol.source)) + ' (use the ⚙ next to Source).');
    if (!cand.length) return result(q, 0, 'No query found. Use Get Data to create one.');
    const needSteps = q.steps || [...new Set((sol.steps || []).map((s) => s.type))];
    let best = null, bestFrac = -1, bestMsg = '';
    for (const def of cand) {
      const have = new Set((def.steps || []).map((s) => s.type));
      const missing = needSteps.filter((s) => ![].concat(s).some((x) => have.has(x)));
      let frac = 1 - (needSteps.length ? missing.length / needSteps.length : 0) * 0.5;
      const label = (m) => (root.DIM && root.DIM.Query ? root.DIM.Query.STEP_LABELS[m] : m) || m;
      let msg = missing.length ? 'add a ' + missing.map((m) => [].concat(m).map(label).join(' or ')).join(', ') + ' step' : '';
      // compare output with the key's output
      const keyDef = key.queries.find((k) => k.name === sol.name) || key.queries[0];
      if (keyDef) {
        const mine = wb.runQuery(def), theirs = key.runQuery(keyDef);
        if (mine.error) { frac -= 0.5; msg = (msg ? msg + '; ' : '') + 'the query has an error: ' + mine.error; }
        else if (!theirs.error) {
          const sameRows = mine.table.rows.length === theirs.table.rows.length;
          if (!sameRows) { frac -= 0.25; msg = (msg ? msg + '; ' : '') + 'it returns ' + mine.table.rows.length + ' rows, expected ' + theirs.table.rows.length; }
        }
      }
      if (q.loaded && !wb.getSheet(def.loadedSheet || def.name)) { frac -= 0.25; msg = (msg ? msg + '; ' : '') + 'use Close & Load so the results land in a sheet'; }
      else if (q.loaded) {
        const now = wb.runQuery(def);
        if (!now.error && def.loadedRows !== undefined && def.loadedRows !== now.table.rows.length) { frac -= 0.25; msg = (msg ? msg + '; ' : '') + 'the loaded sheet is out of date: Close & Load (or Refresh All) again'; }
      }
      if (frac > bestFrac) { bestFrac = frac; best = def; bestMsg = msg; }
    }
    return result(q, bestFrac, bestMsg ? 'Query "' + best.name + '": ' + bestMsg + '.' : 'Query "' + best.name + '" has the right steps and output.');
  }

  function normText(s) { return String(s ?? '').trim().toLowerCase().replace(/[^a-z0-9%.$ -]/g, '').replace(/\s+/g, ' '); }

  function gradeChoice(q, answers) {
    const a = answers[q.id];
    const correct = [].concat(q.answer).sort().join(',');
    if (q.multi) {
      const picked = [].concat(a || []).map(Number).sort();
      const want = [].concat(q.answer).map(Number);
      const hits = picked.filter((x) => want.includes(x)).length;
      const wrongPicks = picked.filter((x) => !want.includes(x)).length;
      const frac = Math.max(0, (hits - wrongPicks) / want.length);
      return result(q, frac, frac === 1 ? 'Correct.' : 'Not quite: ' + hits + ' of ' + want.length + ' correct choices selected' + (wrongPicks ? ', plus ' + wrongPicks + ' incorrect' : '') + '.' + (q.explain ? ' ' + q.explain : ''));
    }
    if (a === undefined || a === null || a === '') return result(q, 0, 'No answer selected.' + (q.explain ? ' ' + q.explain : ''));
    const ok = String(Number(a)) === correct;
    return result(q, ok ? 1 : 0, ok ? 'Correct.' + (q.explainCorrect && q.explain ? ' ' + q.explain : '') : 'Incorrect. The answer is "' + q.options[Number(q.answer)] + '".' + (q.explain ? ' ' + q.explain : ''));
  }

  function gradeNumber(q, answers) {
    const raw = answers[q.id];
    if (raw === undefined || raw === null || String(raw).trim() === '') return result(q, 0, 'No answer entered.');
    const p = typeof raw === 'number' ? { value: raw } : F.parseNumberLike(String(raw));
    if (!p) return result(q, 0, '"' + raw + '" is not a number.');
    const tol = q.tolerance !== undefined ? q.tolerance : 1e-9;
    const ok = Math.abs(p.value - q.answer) <= tol;
    return result(q, ok ? 1 : 0, ok ? 'Correct.' : 'Incorrect. The answer is ' + q.answer + '.' + (q.explain ? ' ' + q.explain : ''));
  }

  function gradeText(q, answers) {
    const a = normText(answers[q.id]);
    if (!a) return result(q, 0, 'No answer entered.');
    const ok = (q.accept || []).some((acc) => (acc instanceof RegExp ? acc.test(String(answers[q.id]).trim()) : normText(acc) === a));
    return result(q, ok ? 1 : 0, ok ? 'Correct.' : 'Not accepted. A correct answer is "' + (q.accept[0] instanceof RegExp ? q.display || 'see notes' : q.accept[0]) + '".' + (q.explain ? ' ' + q.explain : ''));
  }

  function gradeMatch(q, answers) {
    const a = answers[q.id] || {};
    let right = 0;
    q.pairs.forEach(([term, def], i) => { if (a[i] !== undefined && q.pairs[Number(a[i])] && q.pairs[Number(a[i])][1] === def) right++; });
    const frac = right / q.pairs.length;
    return result(q, frac, right === q.pairs.length ? 'All matches are correct.' : right + ' of ' + q.pairs.length + ' matches are correct.');
  }

  // ---------------------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------------------

  /**
   * grade(assessment, { workbook, answers }, opts) -> report
   *   workbook: a Workbook (from the web grid or an uploaded .xlsx), optional
   *   answers:  { [questionId]: value } for non-workbook questions
   *   opts.reveal: include expected values in feedback (after submission)
   */
  function grade(assessment, submission, opts = {}) {
    const answers = submission.answers || {};
    const wb = submission.workbook || buildTemplate(assessment);
    wb.recalc();
    const needsKey = (assessment.questions || []).some(isWorkbookQuestion);
    const key = needsKey ? buildKey(assessment) : null;
    const items = (assessment.questions || []).map((q) => {
      try {
        switch (q.type) {
          case 'cell': return gradeCell(q, wb, key, assessment, opts);
          case 'range': return gradeRange(q, wb, key, assessment, opts);
          case 'cellType': return gradeCellType(q, wb, key, assessment, opts);
          case 'filled': return gradeFilled(q, wb, key, assessment, opts);
          case 'format': return gradeFormat(q, wb, key, assessment, opts);
          case 'table': return gradeTable(q, wb, key, assessment, opts);
          case 'validation': return gradeValidation(q, wb, key, assessment, opts);
          case 'condFormat': return gradeCondFormat(q, wb, key, assessment, opts);
          case 'sorted': return gradeSorted(q, wb, key, assessment, opts);
          case 'chart': return gradeChart(q, wb, key, assessment, opts);
          case 'pivot': return gradePivot(q, wb, key, assessment, opts);
          case 'query': return gradeQuery(q, wb, key, assessment, opts);
          case 'choice': return gradeChoice(q, answers);
          case 'number': return gradeNumber(q, answers);
          case 'text': return gradeText(q, answers);
          case 'match': return gradeMatch(q, answers);
          default: return result(q, 0, 'Unknown question type "' + q.type + '".');
        }
      } catch (e) {
        return result(q, 0, 'Could not grade this item: ' + e.message);
      }
    });
    const max = items.reduce((a, b) => a + b.points, 0);
    const score = Math.round(items.reduce((a, b) => a + b.earned, 0) * 100) / 100;
    return {
      assessmentId: assessment.id,
      title: assessment.title,
      score, max,
      percent: max ? Math.round((score / max) * 1000) / 10 : 0,
      items,
      gradedAt: new Date().toISOString(),
    };
  }

  // Answers that earn full marks on the non-workbook questions (for tests/demos).
  function perfectAnswers(assessment) {
    const out = {};
    for (const q of assessment.questions || []) {
      if (q.type === 'choice') out[q.id] = q.multi ? [].concat(q.answer) : q.answer;
      else if (q.type === 'number') out[q.id] = q.answer;
      else if (q.type === 'text') out[q.id] = q.accept.find((a) => typeof a === 'string');
      else if (q.type === 'match') out[q.id] = Object.fromEntries(q.pairs.map((p, i) => [i, i]));
    }
    return out;
  }

  function letterGrade(pct) {
    if (pct >= 93) return 'A'; if (pct >= 90) return 'A-'; if (pct >= 87) return 'B+'; if (pct >= 83) return 'B';
    if (pct >= 80) return 'B-'; if (pct >= 77) return 'C+'; if (pct >= 73) return 'C'; if (pct >= 70) return 'C-';
    if (pct >= 67) return 'D+'; if (pct >= 60) return 'D'; return 'F';
  }

  const api = { grade, buildKey, buildTemplate, applySolution, perfectAnswers, isWorkbookQuestion, valuesMatch, letterGrade, WORKBOOK_TYPES };
  root.DIM = root.DIM || {};
  root.DIM.Grader = api;
  if (isNode) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
