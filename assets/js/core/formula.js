/*
 * formula.js — a small, dependency-free Excel formula engine.
 *
 * Responsibilities:
 *   - cell address helpers (A1 <-> row/col, ranges, shifting relative refs)
 *   - value coercion that mirrors Excel (numbers, text, booleans, blanks, errors)
 *   - tokenizer + recursive-descent parser producing an AST
 *   - evaluator with ~90 commonly taught worksheet functions
 *
 * The evaluator never touches a workbook directly. It receives a context with
 *   ctx.getCell(sheetName, row, col)  -> evaluated cell value (null for blank)
 *   ctx.sheet                          -> name of the sheet the formula lives in
 *   ctx.row, ctx.col                   -> position of the formula cell
 *   ctx.resolveName(name)              -> range object for a defined name, or null
 *   ctx.resolveTable(table, column, thisRow) -> range object, or null
 *   ctx.usedBounds(sheetName)          -> { rows, cols } for whole-column refs
 *
 * Works in the browser (window.DIM.Formula) and in Node (module.exports).
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------------------
  // Errors and value helpers
  // ---------------------------------------------------------------------------

  class XLError {
    constructor(code, hint) {
      this.code = code; // e.g. '#DIV/0!'
      this.hint = hint || '';
    }
    toString() { return this.code; }
  }

  const ERR = {
    DIV0: '#DIV/0!',
    NA: '#N/A',
    VALUE: '#VALUE!',
    REF: '#REF!',
    NAME: '#NAME?',
    NUM: '#NUM!',
    NULL: '#NULL!',
    CIRC: '#CIRC!',
    SPILL: '#SPILL!',
  };
  const ERROR_CODES = Object.values(ERR);

  function err(code, hint) { return new XLError(code, hint); }
  function isError(v) { return v instanceof XLError; }
  function isBlank(v) { return v === null || v === undefined; }

  // A reference to a rectangular block of cells. Rows/cols are 0-based.
  class RangeRef {
    constructor(sheet, r1, c1, r2, c2) {
      this.sheet = sheet;
      this.r1 = Math.min(r1, r2);
      this.c1 = Math.min(c1, c2);
      this.r2 = Math.max(r1, r2);
      this.c2 = Math.max(c1, c2);
    }
    get rows() { return this.r2 - this.r1 + 1; }
    get cols() { return this.c2 - this.c1 + 1; }
    get size() { return this.rows * this.cols; }
  }
  function isRange(v) { return v instanceof RangeRef; }

  // ---------------------------------------------------------------------------
  // Address helpers
  // ---------------------------------------------------------------------------

  function colToIndex(letters) {
    let n = 0;
    const s = letters.toUpperCase();
    for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  function indexToCol(index) {
    let n = index + 1;
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function toAddress(row, col, colAbs, rowAbs) {
    return (colAbs ? '$' : '') + indexToCol(col) + (rowAbs ? '$' : '') + (row + 1);
  }

  const ADDR_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/;

  function parseAddress(addr) {
    const m = ADDR_RE.exec(String(addr).trim());
    if (!m) return null;
    const row = parseInt(m[4], 10) - 1;
    const col = colToIndex(m[2]);
    if (row < 0 || col < 0 || col > 16383 || row > 1048575) return null;
    return { row, col, colAbs: m[1] === '$', rowAbs: m[3] === '$' };
  }

  // "A1:C3", "B2", "Sheet1!A1:B2" -> { sheet, r1, c1, r2, c2 }
  function parseRangeAddress(text) {
    let s = String(text).trim();
    let sheet = null;
    const bang = s.lastIndexOf('!');
    if (bang >= 0) {
      sheet = s.slice(0, bang);
      if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
      s = s.slice(bang + 1);
    }
    const parts = s.split(':');
    const a = parseAddress(parts[0]);
    const b = parts.length > 1 ? parseAddress(parts[1]) : a;
    if (!a || !b || parts.length > 2) return null;
    return {
      sheet,
      r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col),
      r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col),
    };
  }

  function rangeToAddress(r) {
    const a = toAddress(r.r1, r.c1);
    const b = toAddress(r.r2, r.c2);
    return a === b ? a : a + ':' + b;
  }

  function quoteSheet(name) {
    return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !ADDR_RE.test(name)
      ? name
      : "'" + String(name).replace(/'/g, "''") + "'";
  }

  // ---------------------------------------------------------------------------
  // Dates (Excel 1900 date system, including the famous 1900 leap-year bug)
  // ---------------------------------------------------------------------------

  const MS_PER_DAY = 86400000;
  const EPOCH = Date.UTC(1899, 11, 30); // serial 0 for dates after 1900-03-01

  function dateToSerial(y, m, d) {
    // m is 1-based; JS Date handles overflow (month 13, day 0, ...)
    const t = Date.UTC(y, m - 1, d);
    let serial = Math.round((t - EPOCH) / MS_PER_DAY);
    if (serial < 61) serial -= 1; // before the fake 1900-02-29
    return serial;
  }

  function serialToDate(serial) {
    let s = Math.floor(serial);
    if (s < 61) s += 1;
    const dt = new Date(EPOCH + s * MS_PER_DAY);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay() };
  }

  function todaySerial() {
    const now = new Date();
    return dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  }

  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
    'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  // ---------------------------------------------------------------------------
  // Parsing typed input ("$1,200", "45%", "3/15/2026", "TRUE")
  // ---------------------------------------------------------------------------

  // Returns { value, format } or null when the text is not number-like.
  function parseNumberLike(text) {
    if (typeof text !== 'string') return null;
    const s = text.trim();
    if (s === '') return null;

    // plain numbers, with optional thousands separators and sign
    let m = /^([+-]?)(\d{1,3}(?:,\d{3})+|\d*)(\.\d+)?(?:[eE]([+-]?\d+))?$/.exec(s);
    if (m && (m[2] !== '' || m[3])) {
      const hasComma = m[2].includes(',');
      const v = Number((m[1] + m[2].replace(/,/g, '') + (m[3] || '') + (m[4] ? 'e' + m[4] : '')));
      if (!Number.isNaN(v)) return { value: v, format: hasComma ? (m[3] ? 'number' : 'integer') : null };
    }
    // currency
    m = /^([+-]?)\$\s?(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?$/.exec(s) || /^\(\$(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\)$/.exec(s);
    if (m) {
      if (s.startsWith('(')) return { value: -Number(m[1].replace(/,/g, '') + (m[2] || '')), format: 'currency' };
      return { value: Number(m[1] + m[2].replace(/,/g, '') + (m[3] || '')), format: 'currency' };
    }
    // percent
    m = /^([+-]?\d*\.?\d+)\s?%$/.exec(s);
    if (m) {
      const decimals = (m[1].split('.')[1] || '').length;
      return { value: round15(Number(m[1]) / 100), format: decimals > 0 ? 'percent1' : 'percent' };
    }
    // dates m/d/yyyy, m/d/yy, yyyy-mm-dd, m-d-yyyy
    m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})$/.exec(s);
    if (m) {
      let y = parseInt(m[3], 10);
      if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
      const mo = parseInt(m[1], 10), d = parseInt(m[2], 10);
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31 && validDay(y, mo, d)) return { value: dateToSerial(y, mo, d), format: 'date' };
    }
    m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (m) {
      const y = +m[1], mo = +m[2], d = +m[3];
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31 && validDay(y, mo, d)) return { value: dateToSerial(y, mo, d), format: 'date' };
    }
    return null;
  }

  function validDay(y, m, d) {
    const dt = new Date(Date.UTC(y, m - 1, d));
    return dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
  }

  function round15(x) {
    if (!Number.isFinite(x) || x === 0) return x;
    return Number(x.toPrecision(15));
  }

  // ---------------------------------------------------------------------------
  // Coercion
  // ---------------------------------------------------------------------------

  function toNumber(v) {
    if (isError(v)) return v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isBlank(v)) return 0;
    if (typeof v === 'string') {
      if (v.trim() === '') return err(ERR.VALUE, 'Text cannot be used as a number');
      const p = parseNumberLike(v);
      if (p) return p.value;
      return err(ERR.VALUE, '"' + v + '" is text, not a number');
    }
    return err(ERR.VALUE);
  }

  function toText(v) {
    if (isError(v)) return v;
    if (isBlank(v)) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return formatGeneral(v);
    return String(v);
  }

  function toBool(v) {
    if (isError(v)) return v;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (isBlank(v)) return false;
    if (typeof v === 'string') {
      const u = v.trim().toUpperCase();
      if (u === 'TRUE') return true;
      if (u === 'FALSE') return false;
      return err(ERR.VALUE, 'Expected TRUE or FALSE');
    }
    return err(ERR.VALUE);
  }

  // Excel ordering for comparisons: numbers < text < booleans; text is case-insensitive.
  function typeRank(v) {
    if (typeof v === 'number') return 0;
    if (typeof v === 'string') return 1;
    if (typeof v === 'boolean') return 2;
    return 0;
  }

  function compare(a, b) {
    if (isBlank(a) && isBlank(b)) return 0;
    if (isBlank(a)) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (isBlank(b)) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = typeRank(a), rb = typeRank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (typeof a === 'string') {
      const x = a.toLowerCase(), y = b.toLowerCase();
      return x < y ? -1 : x > y ? 1 : 0;
    }
    if (typeof a === 'boolean') return a === b ? 0 : a ? 1 : -1;
    const d = round15(a - b);
    return d < 0 ? -1 : d > 0 ? 1 : 0;
  }

  // ---------------------------------------------------------------------------
  // Number formatting (shared by cell display and TEXT())
  // ---------------------------------------------------------------------------

  function formatGeneral(n) {
    if (!Number.isFinite(n)) return String(n);
    if (n === 0) return '0';
    const abs = Math.abs(n);
    if (abs >= 1e11 || abs < 1e-9) {
      return n.toExponential(5).replace(/\.?0+e/, 'E').replace('E+', 'E+').replace(/e/, 'E');
    }
    let s = Number(n.toPrecision(10)).toString();
    if (s.includes('e')) s = Number(n.toPrecision(10)).toFixed(10).replace(/\.?0+$/, '');
    return s;
  }

  function addThousands(intStr) {
    return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  function formatFixed(n, decimals, thousands) {
    const neg = n < 0;
    const fixed = Math.abs(n).toFixed(decimals);
    let [i, f] = fixed.split('.');
    if (thousands) i = addThousands(i);
    return (neg && Number(fixed) !== 0 ? '-' : '') + i + (f !== undefined ? '.' + f : '');
  }

  function formatDatePattern(serial, pattern) {
    const { y, m, d, dow } = serialToDate(serial);
    const pad = (x) => String(x).padStart(2, '0');
    return pattern.replace(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d/gi, (tok) => {
      switch (tok.toLowerCase()) {
        case 'yyyy': return String(y);
        case 'yy': return pad(y % 100);
        case 'mmmm': return MONTHS[m - 1];
        case 'mmm': return MONTHS[m - 1].slice(0, 3);
        case 'mm': return pad(m);
        case 'm': return String(m);
        case 'dddd': return DAYS[dow];
        case 'ddd': return DAYS[dow].slice(0, 3);
        case 'dd': return pad(d);
        case 'd': return String(d);
        default: return tok;
      }
    });
  }

  // Applies an Excel-style number format code (subset) to a number.
  function applyFormatCode(n, code) {
    const c = String(code);
    if (/[yd]/i.test(c) || /^m{1,4}([^0#]*)$/i.test(c)) return formatDatePattern(n, c);
    let s = c;
    let prefix = '', suffix = '';
    const pm = /^([^0#,.]*)/.exec(s);
    if (pm) { prefix = pm[1]; s = s.slice(prefix.length); }
    const sm = /([^0#,.]*)$/.exec(s);
    if (sm) { suffix = sm[1]; s = s.slice(0, s.length - suffix.length); }
    let value = n;
    if (suffix.includes('%')) value = n * 100;
    const decimals = s.includes('.') ? s.split('.')[1].replace(/[^0#]/g, '').length : 0;
    const thousands = s.includes(',');
    let out = formatFixed(value, decimals, thousands);
    if (out.startsWith('-') && prefix) { out = out.slice(1); prefix = '-' + prefix; }
    return prefix.replace(/"/g, '') + out + suffix.replace(/"/g, '');
  }

  const NAMED_FORMATS = {
    general: null,
    number: '#,##0.00',
    integer: '#,##0',
    currency: '$#,##0.00',
    percent: '0%',
    percent1: '0.0%',
    date: 'm/d/yyyy',
    longdate: 'mmmm d, yyyy',
    text: '@',
  };

  // Display a value the way a cell shows it, given a named format.
  function formatValue(v, format) {
    if (isError(v)) return v.code;
    if (isBlank(v)) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') {
      const code = format && NAMED_FORMATS[format] !== undefined ? NAMED_FORMATS[format] : format;
      if (!code || code === '@' || format === 'general') return formatGeneral(v);
      if (format === 'currency' && v < 0) return '-' + applyFormatCode(-v, code);
      return applyFormatCode(v, code);
    }
    return String(v);
  }

  // ---------------------------------------------------------------------------
  // Tokenizer
  // ---------------------------------------------------------------------------

  const T = {
    NUM: 'num', STR: 'str', BOOL: 'bool', ERR: 'err', REF: 'ref', COLS: 'cols', ROWS: 'rows',
    FUNC: 'func', NAME: 'name', TABLE: 'table', OP: 'op', LP: '(', RP: ')', COMMA: ',', WS: 'ws',
  };

  const SHEET_PREFIX = "(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!";
  const REF_RE = new RegExp('^(?:' + SHEET_PREFIX + ')?(\\$?[A-Za-z]{1,3}\\$?\\d+)(?![A-Za-z0-9_(.\\[])');
  const COLS_RE = new RegExp('^(?:' + SHEET_PREFIX + ')?(\\$?[A-Za-z]{1,3}):(\\$?[A-Za-z]{1,3})(?![A-Za-z0-9_(])');
  const ROWS_RE = new RegExp('^(?:' + SHEET_PREFIX + ')?(\\$?\\d+):(\\$?\\d+)(?![A-Za-z0-9_(])');

  function tokenize(src) {
    const tokens = [];
    let i = 0;
    const n = src.length;
    while (i < n) {
      const ch = src[i];
      const rest = src.slice(i);
      let m;

      if (/\s/.test(ch)) {
        let j = i;
        while (j < n && /\s/.test(src[j])) j++;
        tokens.push({ type: T.WS, text: src.slice(i, j), pos: i });
        i = j;
        continue;
      }
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < n) {
          if (src[j] === '"') {
            if (src[j + 1] === '"') { s += '"'; j += 2; continue; }
            break;
          }
          s += src[j++];
        }
        if (j >= n) throw new SyntaxError('Missing closing quote');
        tokens.push({ type: T.STR, value: s, text: src.slice(i, j + 1), pos: i });
        i = j + 1;
        continue;
      }
      if (ch === '#') {
        const code = ERROR_CODES.find((c) => rest.toUpperCase().startsWith(c));
        if (code) {
          tokens.push({ type: T.ERR, value: code, text: rest.slice(0, code.length), pos: i });
          i += code.length;
          continue;
        }
      }
      // whole column / whole row references (A:A, $B:$D, 2:5)
      if ((m = COLS_RE.exec(rest))) {
        tokens.push({ type: T.COLS, sheet: m[1] ? m[1].replace(/''/g, "'") : m[2] || null, a: m[3], b: m[4], text: m[0], pos: i });
        i += m[0].length;
        continue;
      }
      if ((m = REF_RE.exec(rest))) {
        tokens.push({ type: T.REF, sheet: m[1] ? m[1].replace(/''/g, "'") : m[2] || null, addr: m[3], text: m[0], pos: i });
        i += m[0].length;
        continue;
      }
      if (/[0-9.]/.test(ch)) {
        if ((m = ROWS_RE.exec(rest)) && !/^\d+\.\d/.test(rest)) {
          tokens.push({ type: T.ROWS, sheet: null, a: m[3], b: m[4], text: m[0], pos: i });
          i += m[0].length;
          continue;
        }
        m = /^(\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);
        if (m) {
          tokens.push({ type: T.NUM, value: Number(m[0]), text: m[0], pos: i });
          i += m[0].length;
          continue;
        }
      }
      // structured reference with no table name: [@Column] or [Column]
      if (ch === '[') {
        const end = findBracketEnd(src, i);
        tokens.push({ type: T.TABLE, table: null, spec: src.slice(i + 1, end), text: src.slice(i, end + 1), pos: i });
        i = end + 1;
        continue;
      }
      if (/[A-Za-z_\\]/.test(ch)) {
        m = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(rest);
        const word = m[0];
        let j = i + word.length;
        if (src[j] === '(') {
          tokens.push({ type: T.FUNC, name: word.toUpperCase().replace(/^_XLFN\./, ''), text: word, pos: i });
          i = j;
          continue;
        }
        if (src[j] === '[') {
          const end = findBracketEnd(src, j);
          tokens.push({ type: T.TABLE, table: word, spec: src.slice(j + 1, end), text: src.slice(i, end + 1), pos: i });
          i = end + 1;
          continue;
        }
        const up = word.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') {
          tokens.push({ type: T.BOOL, value: up === 'TRUE', text: word, pos: i });
        } else {
          tokens.push({ type: T.NAME, name: word, text: word, pos: i });
        }
        i = j;
        continue;
      }
      const two = src.slice(i, i + 2);
      if (two === '<=' || two === '>=' || two === '<>') {
        tokens.push({ type: T.OP, value: two, text: two, pos: i });
        i += 2;
        continue;
      }
      if ('+-*/^&=<>%:'.includes(ch)) {
        tokens.push({ type: T.OP, value: ch, text: ch, pos: i });
        i++;
        continue;
      }
      if (ch === '(') { tokens.push({ type: T.LP, text: ch, pos: i }); i++; continue; }
      if (ch === ')') { tokens.push({ type: T.RP, text: ch, pos: i }); i++; continue; }
      if (ch === ',' || ch === ';') { tokens.push({ type: T.COMMA, text: ch, pos: i }); i++; continue; }
      throw new SyntaxError('Unexpected character "' + ch + '"');
    }
    return tokens;
  }

  function findBracketEnd(src, start) {
    let depth = 0;
    for (let j = start; j < src.length; j++) {
      if (src[j] === '[') depth++;
      else if (src[j] === ']') { depth--; if (depth === 0) return j; }
    }
    throw new SyntaxError('Missing closing ]');
  }

  // ---------------------------------------------------------------------------
  // Parser -> AST
  // ---------------------------------------------------------------------------

  function parse(formula) {
    let src = String(formula);
    if (src.startsWith('=')) src = src.slice(1);
    const tokens = tokenize(src).filter((t) => t.type !== T.WS);
    let p = 0;
    const peek = () => tokens[p];
    const next = () => tokens[p++];
    const isOp = (t, ...ops) => t && t.type === T.OP && ops.includes(t.value);

    function expression() { return comparison(); }

    function comparison() {
      let left = concat();
      while (isOp(peek(), '=', '<>', '<', '>', '<=', '>=')) {
        const op = next().value;
        left = { type: 'binary', op, left, right: concat() };
      }
      return left;
    }
    function concat() {
      let left = additive();
      while (isOp(peek(), '&')) { next(); left = { type: 'binary', op: '&', left, right: additive() }; }
      return left;
    }
    function additive() {
      let left = multiplicative();
      while (isOp(peek(), '+', '-')) {
        const op = next().value;
        left = { type: 'binary', op, left, right: multiplicative() };
      }
      return left;
    }
    function multiplicative() {
      let left = power();
      while (isOp(peek(), '*', '/')) {
        const op = next().value;
        left = { type: 'binary', op, left, right: power() };
      }
      return left;
    }
    function power() {
      let left = unary();
      while (isOp(peek(), '^')) { next(); left = { type: 'binary', op: '^', left, right: unary() }; }
      return left;
    }
    function unary() {
      if (isOp(peek(), '-', '+')) {
        const op = next().value;
        return { type: 'unary', op, operand: unary() };
      }
      return percent();
    }
    function percent() {
      let node = range();
      while (isOp(peek(), '%')) { next(); node = { type: 'percent', operand: node }; }
      return node;
    }
    function range() {
      let left = primary();
      while (isOp(peek(), ':')) {
        next();
        const right = primary();
        if (left.type !== 'ref' || right.type !== 'ref') throw new SyntaxError('A range needs two cell addresses, like A1:B5');
        left = { type: 'range', start: left, end: right, sheet: left.sheet };
      }
      return left;
    }
    function primary() {
      const t = next();
      if (!t) throw new SyntaxError('The formula ends too early');
      switch (t.type) {
        case T.NUM: return { type: 'num', value: t.value };
        case T.STR: return { type: 'str', value: t.value };
        case T.BOOL: return { type: 'bool', value: t.value };
        case T.ERR: return { type: 'err', value: t.value };
        case T.REF: {
          const a = parseAddress(t.addr);
          if (!a) return { type: 'err', value: ERR.REF };
          return { type: 'ref', sheet: t.sheet, ...a };
        }
        case T.COLS: {
          const a = colToIndex(t.a.replace('$', '')), b = colToIndex(t.b.replace('$', ''));
          return { type: 'cols', sheet: t.sheet, c1: Math.min(a, b), c2: Math.max(a, b) };
        }
        case T.ROWS: {
          const a = parseInt(t.a.replace('$', ''), 10) - 1, b = parseInt(t.b.replace('$', ''), 10) - 1;
          return { type: 'rows', sheet: t.sheet, r1: Math.min(a, b), r2: Math.max(a, b) };
        }
        case T.TABLE: return { type: 'table', table: t.table, spec: t.spec };
        case T.NAME: return { type: 'name', name: t.name };
        case T.FUNC: {
          const args = [];
          const lp = next();
          if (!lp || lp.type !== T.LP) throw new SyntaxError('Expected ( after ' + t.name);
          if (peek() && peek().type === T.RP) { next(); return { type: 'func', name: t.name, args }; }
          for (;;) {
            if (peek() && (peek().type === T.COMMA || peek().type === T.RP)) args.push({ type: 'missing' });
            else args.push(expression());
            const sep = next();
            if (!sep) throw new SyntaxError('Missing ) to close ' + t.name + '(');
            if (sep.type === T.RP) break;
            if (sep.type !== T.COMMA) throw new SyntaxError('Expected , or ) in ' + t.name + '(');
          }
          return { type: 'func', name: t.name, args };
        }
        case T.LP: {
          const e = expression();
          const rp = next();
          if (!rp || rp.type !== T.RP) throw new SyntaxError('Missing )');
          return e;
        }
        case T.OP:
          throw new SyntaxError('Unexpected "' + t.value + '"');
        default:
          throw new SyntaxError('Unexpected "' + t.text + '"');
      }
    }

    if (tokens.length === 0) throw new SyntaxError('The formula is empty');
    const ast = expression();
    if (p < tokens.length) throw new SyntaxError('Unexpected "' + tokens[p].text + '"');
    return ast;
  }

  // Walk an AST and collect function names (used by the grader).
  function functionsUsed(ast, out = new Set()) {
    if (!ast || typeof ast !== 'object') return out;
    if (ast.type === 'func') out.add(ast.name);
    for (const k of ['left', 'right', 'operand', 'start', 'end']) if (ast[k]) functionsUsed(ast[k], out);
    if (ast.args) ast.args.forEach((a) => functionsUsed(a, out));
    return out;
  }

  // Shift the relative references in a formula by (dRow, dCol), as Excel does on fill/copy.
  function shiftFormula(formula, dRow, dCol) {
    if (!String(formula).startsWith('=')) return formula;
    const body = String(formula).slice(1);
    let tokens;
    try { tokens = tokenize(body); } catch (e) { return formula; }
    let out = '=';
    for (const t of tokens) {
      if (t.type === T.REF) {
        const a = parseAddress(t.addr);
        const row = a.rowAbs ? a.row : a.row + dRow;
        const col = a.colAbs ? a.col : a.col + dCol;
        const prefix = t.sheet ? quoteSheet(t.sheet) + '!' : '';
        out += row < 0 || col < 0 ? '#REF!' : prefix + toAddress(row, col, a.colAbs, a.rowAbs);
      } else if (t.type === T.COLS) {
        const shiftCol = (s) => (s.startsWith('$') ? s : indexToCol(Math.max(0, colToIndex(s) + dCol)));
        out += (t.sheet ? quoteSheet(t.sheet) + '!' : '') + shiftCol(t.a) + ':' + shiftCol(t.b);
      } else {
        out += t.text;
      }
    }
    return out;
  }

  // List the references (as ranges) a formula reads. Used for highlighting.
  function referencedRanges(formula) {
    const body = String(formula).replace(/^=/, '');
    let tokens;
    try { tokens = tokenize(body); } catch (e) { return []; }
    const out = [];
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      if (t.type !== T.REF) continue;
      const a = parseAddress(t.addr);
      let j = i + 1;
      while (tokens[j] && tokens[j].type === T.WS) j++;
      if (tokens[j] && tokens[j].type === T.OP && tokens[j].value === ':') {
        let k = j + 1;
        while (tokens[k] && tokens[k].type === T.WS) k++;
        if (tokens[k] && tokens[k].type === T.REF) {
          const b = parseAddress(tokens[k].addr);
          out.push({ sheet: t.sheet, r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) });
          i = k;
          continue;
        }
      }
      out.push({ sheet: t.sheet, r1: a.row, c1: a.col, r2: a.row, c2: a.col });
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // Evaluator
  // ---------------------------------------------------------------------------

  function evaluate(ast, ctx) {
    switch (ast.type) {
      case 'num': return ast.value;
      case 'str': return ast.value;
      case 'bool': return ast.value;
      case 'err': return err(ast.value);
      case 'missing': return null;
      case 'ref':
        return new RangeRef(ast.sheet || ctx.sheet, ast.row, ast.col, ast.row, ast.col);
      case 'range': {
        const s = ast.start, e = ast.end;
        return new RangeRef(s.sheet || ctx.sheet, s.row, s.col, e.row, e.col);
      }
      case 'cols': {
        const sheet = ast.sheet || ctx.sheet;
        const b = ctx.usedBounds ? ctx.usedBounds(sheet) : { rows: 1000 };
        return new RangeRef(sheet, 0, ast.c1, Math.max(0, b.rows - 1), ast.c2);
      }
      case 'rows': {
        const sheet = ast.sheet || ctx.sheet;
        const b = ctx.usedBounds ? ctx.usedBounds(sheet) : { cols: 26 };
        return new RangeRef(sheet, ast.r1, 0, ast.r2, Math.max(0, b.cols - 1));
      }
      case 'name': {
        const r = ctx.resolveName ? ctx.resolveName(ast.name) : null;
        if (r) return r;
        return err(ERR.NAME, 'Unknown name "' + ast.name + '"');
      }
      case 'table': {
        const r = ctx.resolveTable ? ctx.resolveTable(ast.table, ast.spec, ctx) : null;
        if (r) return r;
        return err(ERR.REF, 'Unknown table reference');
      }
      case 'unary': {
        const v = toNumber(deref(evaluate(ast.operand, ctx), ctx));
        if (isError(v)) return v;
        return ast.op === '-' ? -v : v;
      }
      case 'percent': {
        const v = toNumber(deref(evaluate(ast.operand, ctx), ctx));
        if (isError(v)) return v;
        return round15(v / 100);
      }
      case 'binary': return evalBinary(ast, ctx);
      case 'func': return evalFunc(ast, ctx);
      default: return err(ERR.VALUE);
    }
  }

  // Turn a single-cell range into its value; multi-cell ranges in a scalar
  // position use implicit intersection with the formula's row/column.
  function deref(v, ctx) {
    if (!isRange(v)) return v;
    if (v.size === 1) return ctx.getCell(v.sheet, v.r1, v.c1);
    if (v.cols === 1 && ctx.row >= v.r1 && ctx.row <= v.r2 && v.sheet === ctx.sheet) return ctx.getCell(v.sheet, ctx.row, v.c1);
    if (v.rows === 1 && ctx.col >= v.c1 && ctx.col <= v.c2 && v.sheet === ctx.sheet) return ctx.getCell(v.sheet, v.r1, ctx.col);
    return err(ERR.VALUE, 'A range was used where one value was expected');
  }

  function evalBinary(ast, ctx) {
    const a = deref(evaluate(ast.left, ctx), ctx);
    const b = deref(evaluate(ast.right, ctx), ctx);
    if (isError(a)) return a;
    if (isError(b)) return b;
    const op = ast.op;
    if (op === '&') return toText(a) + toText(b);
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const c = compare(a, b);
      switch (op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '<': return c < 0;
        case '>': return c > 0;
        case '<=': return c <= 0;
        case '>=': return c >= 0;
      }
    }
    const x = toNumber(a), y = toNumber(b);
    if (isError(x)) return x;
    if (isError(y)) return y;
    switch (op) {
      case '+': return round15(x + y);
      case '-': return round15(x - y);
      case '*': return round15(x * y);
      case '/': return y === 0 ? err(ERR.DIV0, 'Dividing by zero (or by a blank cell)') : round15(x / y);
      case '^': {
        const r = Math.pow(x, y);
        return Number.isFinite(r) ? round15(r) : err(ERR.NUM);
      }
    }
    return err(ERR.VALUE);
  }

  // Iterate every value in a list of evaluated arguments.
  // For ranges, yields { value, fromRange: true }; for scalars, fromRange: false.
  function* iterArgs(args, ctx) {
    for (const a of args) {
      if (isRange(a)) {
        for (let r = a.r1; r <= a.r2; r++) {
          for (let c = a.c1; c <= a.c2; c++) yield { value: ctx.getCell(a.sheet, r, c), fromRange: true };
        }
      } else if (Array.isArray(a)) {
        for (const v of a.flat()) yield { value: v, fromRange: true };
      } else {
        yield { value: a, fromRange: false };
      }
    }
  }

  // Numbers for SUM/AVERAGE-style functions: ranges ignore text/bools/blanks,
  // direct arguments are coerced. Errors propagate.
  function collectNumbers(args, ctx) {
    const nums = [];
    for (const { value, fromRange } of iterArgs(args, ctx)) {
      if (isError(value)) return value;
      if (fromRange) {
        if (typeof value === 'number') nums.push(value);
      } else {
        if (isBlank(value)) continue;
        const n = toNumber(value);
        if (isError(n)) return n;
        nums.push(n);
      }
    }
    return nums;
  }

  function rangeMatrix(r, ctx) {
    if (!isRange(r)) return [[r]];
    const rows = [];
    for (let i = r.r1; i <= r.r2; i++) {
      const row = [];
      for (let j = r.c1; j <= r.c2; j++) row.push(ctx.getCell(r.sheet, i, j));
      rows.push(row);
    }
    return rows;
  }

  function rangeList(r, ctx) { return rangeMatrix(r, ctx).flat(); }

  // --- criteria (COUNTIF, SUMIFS, ...) --------------------------------------

  function wildcardToRegex(pattern) {
    let re = '';
    for (let i = 0; i < pattern.length; i++) {
      const ch = pattern[i];
      if (ch === '~' && i + 1 < pattern.length) { re += escapeRe(pattern[++i]); continue; }
      if (ch === '*') re += '.*';
      else if (ch === '?') re += '.';
      else re += escapeRe(ch);
    }
    return new RegExp('^' + re + '$', 'i');
  }
  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  function makeCriterion(crit) {
    if (isError(crit)) return () => false;
    if (typeof crit === 'number') return (v) => typeof v === 'number' ? compare(v, crit) === 0 : (typeof v === 'string' && parseNumberLike(v) && parseNumberLike(v).value === crit);
    if (typeof crit === 'boolean') return (v) => v === crit;
    if (isBlank(crit)) return (v) => isBlank(v) || v === '';
    const s = String(crit);
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
    const op = m[1] || '=';
    const operand = m[2];
    const num = parseNumberLike(operand);
    const upper = operand.toUpperCase();
    if (operand === '') {
      if (op === '=') return (v) => isBlank(v) || v === '';
      if (op === '<>') return (v) => !(isBlank(v) || v === '');
    }
    if (num) {
      const target = num.value;
      return (v) => {
        if (typeof v !== 'number') {
          if (op === '<>') return true;
          if (op === '=' && typeof v === 'string') { const p = parseNumberLike(v); return !!p && p.value === target; }
          return false;
        }
        const c = compare(v, target);
        return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
      };
    }
    if (upper === 'TRUE' || upper === 'FALSE') {
      const b = upper === 'TRUE';
      return (v) => (op === '<>' ? v !== b : op === '=' ? v === b : false);
    }
    if (op === '=' || op === '<>') {
      const re = /[*?~]/.test(operand) ? wildcardToRegex(operand) : null;
      const test = (v) => {
        if (typeof v !== 'string') return false;
        return re ? re.test(v) : v.toLowerCase() === operand.toLowerCase();
      };
      return op === '=' ? test : (v) => !test(v);
    }
    return (v) => {
      if (typeof v !== 'string') return false;
      const c = compare(v, operand);
      return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    };
  }

  function criteriaMask(pairs, ctx, shape) {
    // pairs: [[range, criterion], ...]; every range must match `shape`
    let mask = null;
    for (const [rng, crit] of pairs) {
      if (!isRange(rng)) return err(ERR.VALUE, 'Criteria ranges must be ranges');
      if (shape && (rng.rows !== shape.rows || rng.cols !== shape.cols)) return err(ERR.VALUE, 'All ranges must be the same size');
      const vals = rangeList(rng, ctx);
      const test = makeCriterion(crit);
      const m = vals.map((v) => test(v));
      mask = mask ? mask.map((x, i) => x && m[i]) : m;
    }
    return mask;
  }

  // --- lookup helpers --------------------------------------------------------

  function lookupEquals(a, b, wildcard) {
    if (isBlank(a) || isBlank(b)) return isBlank(a) && isBlank(b);
    if (typeof a !== typeof b) return false; // 5 and "5" do not match (a classic key bug)
    if (typeof a === 'string') {
      if (wildcard && /[*?~]/.test(b)) return wildcardToRegex(b).test(a);
      return a.toLowerCase() === b.toLowerCase();
    }
    return compare(a, b) === 0;
  }

  function xmatchIndex(lookup, list, matchMode, searchMode) {
    const idxs = list.map((_, i) => i);
    if (searchMode === -1) idxs.reverse();
    if (matchMode === 0 || matchMode === 2) {
      for (const i of idxs) if (lookupEquals(list[i], lookup, matchMode === 2)) return i;
      return -1;
    }
    // -1: exact or next smaller, 1: exact or next larger
    let best = -1;
    for (const i of idxs) {
      const v = list[i];
      if (isBlank(v) || typeof v !== typeof lookup) continue;
      const c = compare(v, lookup);
      if (c === 0) return i;
      if (matchMode === -1 && c < 0 && (best < 0 || compare(v, list[best]) > 0)) best = i;
      if (matchMode === 1 && c > 0 && (best < 0 || compare(v, list[best]) < 0)) best = i;
    }
    return best;
  }

  function approxMatchAscending(list, lookup) {
    let best = -1;
    for (let i = 0; i < list.length; i++) {
      const v = list[i];
      if (isBlank(v) || typeof v !== typeof lookup) continue;
      if (compare(v, lookup) <= 0) best = i;
      else break;
    }
    return best;
  }

  // --- function table --------------------------------------------------------

  const LAZY = new Set(['IF', 'IFERROR', 'IFNA', 'IFS', 'AND', 'OR', 'SWITCH', 'CHOOSE']);

  function evalFunc(ast, ctx) {
    const name = ast.name;
    const fn = FUNCTIONS[name];
    if (!fn) return err(ERR.NAME, name + ' is not available in the web workbook. Try it in Excel.');
    if (LAZY.has(name)) return fn(ast.args, ctx);
    const args = ast.args.map((a) => evaluate(a, ctx));
    return fn(args, ctx, ast.args);
  }

  // helpers for arg handling
  const scalar = (v, ctx) => deref(v, ctx);
  const num = (v, ctx) => toNumber(deref(v, ctx));
  const txt = (v, ctx) => toText(deref(v, ctx));
  function argCount(args, min, max) {
    if (args.length < min || (max !== undefined && args.length > max)) return err(ERR.VALUE, 'Wrong number of arguments');
    return null;
  }
  function firstError(...vals) { return vals.find(isError) || null; }

  function aggregate(fnName, args, ctx) {
    const nums = collectNumbers(args, ctx);
    if (isError(nums)) return nums;
    switch (fnName) {
      case 'SUM': return round15(nums.reduce((a, b) => a + b, 0));
      case 'AVERAGE': return nums.length ? round15(nums.reduce((a, b) => a + b, 0) / nums.length) : err(ERR.DIV0, 'No numbers to average');
      case 'MIN': return nums.length ? Math.min(...nums) : 0;
      case 'MAX': return nums.length ? Math.max(...nums) : 0;
      case 'COUNT': return nums.length;
      case 'MEDIAN': {
        if (!nums.length) return err(ERR.NUM);
        const s = [...nums].sort((a, b) => a - b);
        const mid = Math.floor(s.length / 2);
        return s.length % 2 ? s[mid] : round15((s[mid - 1] + s[mid]) / 2);
      }
      case 'PRODUCT': return round15(nums.reduce((a, b) => a * b, 1));
      case 'STDEV': {
        if (nums.length < 2) return err(ERR.DIV0);
        const mean = nums.reduce((a, b) => a + b, 0) / nums.length;
        return round15(Math.sqrt(nums.reduce((a, b) => a + (b - mean) ** 2, 0) / (nums.length - 1)));
      }
    }
    return err(ERR.VALUE);
  }

  function roundTo(n, digits, mode) {
    const f = Math.pow(10, digits);
    const x = round15(Math.abs(n) * f);
    let r;
    if (mode === 'up') r = Math.ceil(x);
    else if (mode === 'down') r = Math.floor(x);
    else r = Math.round(x);
    return round15((Math.sign(n) * r) / f);
  }

  const FUNCTIONS = {
    // ---- math & statistics
    SUM: (a, c) => aggregate('SUM', a, c),
    AVERAGE: (a, c) => aggregate('AVERAGE', a, c),
    MIN: (a, c) => aggregate('MIN', a, c),
    MAX: (a, c) => aggregate('MAX', a, c),
    COUNT: (a, c) => aggregate('COUNT', a, c),
    MEDIAN: (a, c) => aggregate('MEDIAN', a, c),
    PRODUCT: (a, c) => aggregate('PRODUCT', a, c),
    STDEV: (a, c) => aggregate('STDEV', a, c),
    'STDEV.S': (a, c) => aggregate('STDEV', a, c),
    COUNTA: (args, ctx) => {
      let n = 0;
      for (const { value } of iterArgs(args, ctx)) if (!isBlank(value)) n++;
      return n;
    },
    COUNTBLANK: (args, ctx) => {
      const e = argCount(args, 1, 1); if (e) return e;
      if (!isRange(args[0])) return err(ERR.VALUE, 'COUNTBLANK needs a range');
      return rangeList(args[0], ctx).filter((v) => isBlank(v) || v === '').length;
    },
    ROUND: (args, ctx) => {
      const e = argCount(args, 1, 2); if (e) return e;
      const n = num(args[0], ctx), d = args[1] === undefined ? 0 : num(args[1], ctx);
      return firstError(n, d) || roundTo(n, Math.trunc(d));
    },
    ROUNDUP: (args, ctx) => {
      const n = num(args[0], ctx), d = args[1] === undefined ? 0 : num(args[1], ctx);
      return firstError(n, d) || roundTo(n, Math.trunc(d), 'up');
    },
    ROUNDDOWN: (args, ctx) => {
      const n = num(args[0], ctx), d = args[1] === undefined ? 0 : num(args[1], ctx);
      return firstError(n, d) || roundTo(n, Math.trunc(d), 'down');
    },
    INT: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : Math.floor(n); },
    ABS: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : Math.abs(n); },
    SQRT: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : n < 0 ? err(ERR.NUM) : round15(Math.sqrt(n)); },
    POWER: (args, ctx) => { const a = num(args[0], ctx), b = num(args[1], ctx); return firstError(a, b) || round15(Math.pow(a, b)); },
    MOD: (args, ctx) => {
      const a = num(args[0], ctx), b = num(args[1], ctx);
      if (firstError(a, b)) return firstError(a, b);
      if (b === 0) return err(ERR.DIV0);
      return round15(a - b * Math.floor(a / b));
    },
    LARGE: (args, ctx) => {
      const nums = collectNumbers([args[0]], ctx); const k = num(args[1], ctx);
      if (firstError(nums, k)) return firstError(nums, k);
      const s = [...nums].sort((x, y) => y - x);
      return k >= 1 && k <= s.length ? s[Math.ceil(k) - 1] : err(ERR.NUM);
    },
    SMALL: (args, ctx) => {
      const nums = collectNumbers([args[0]], ctx); const k = num(args[1], ctx);
      if (firstError(nums, k)) return firstError(nums, k);
      const s = [...nums].sort((x, y) => x - y);
      return k >= 1 && k <= s.length ? s[Math.ceil(k) - 1] : err(ERR.NUM);
    },
    RANK: (args, ctx) => FUNCTIONS['RANK.EQ'](args, ctx),
    'RANK.EQ': (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      const nums = collectNumbers([args[1]], ctx);
      if (isError(nums)) return nums;
      const asc = args[2] !== undefined && num(args[2], ctx) !== 0;
      if (!nums.includes(n)) return err(ERR.NA);
      return 1 + nums.filter((x) => (asc ? x < n : x > n)).length;
    },
    SUMPRODUCT: (args, ctx) => {
      const lists = args.map((a) => rangeList(a, ctx));
      const len = lists[0].length;
      if (lists.some((l) => l.length !== len)) return err(ERR.VALUE, 'All ranges must be the same size');
      let total = 0;
      for (let i = 0; i < len; i++) {
        let p = 1;
        for (const l of lists) {
          const v = l[i];
          if (isError(v)) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        total += p;
      }
      return round15(total);
    },

    // ---- conditional aggregation
    COUNTIF: (args, ctx) => {
      const e = argCount(args, 2, 2); if (e) return e;
      const mask = criteriaMask([[args[0], scalar(args[1], ctx)]], ctx);
      return isError(mask) ? mask : mask.filter(Boolean).length;
    },
    COUNTIFS: (args, ctx) => {
      if (args.length < 2 || args.length % 2) return err(ERR.VALUE, 'COUNTIFS needs range/criteria pairs');
      const pairs = [];
      for (let i = 0; i < args.length; i += 2) pairs.push([args[i], scalar(args[i + 1], ctx)]);
      const mask = criteriaMask(pairs, ctx, args[0]);
      return isError(mask) ? mask : mask.filter(Boolean).length;
    },
    SUMIF: (args, ctx) => {
      const e = argCount(args, 2, 3); if (e) return e;
      const sumRange = args[2] || args[0];
      const mask = criteriaMask([[args[0], scalar(args[1], ctx)]], ctx);
      if (isError(mask)) return mask;
      if (!isRange(sumRange)) return err(ERR.VALUE);
      const sr = new RangeRef(sumRange.sheet, sumRange.r1, sumRange.c1, sumRange.r1 + args[0].rows - 1, sumRange.c1 + args[0].cols - 1);
      const vals = rangeList(sr, ctx);
      let t = 0;
      mask.forEach((m, i) => { if (m && typeof vals[i] === 'number') t += vals[i]; });
      return round15(t);
    },
    SUMIFS: (args, ctx) => {
      if (args.length < 3 || args.length % 2 === 0) return err(ERR.VALUE, 'SUMIFS(sum_range, criteria_range1, criteria1, ...)');
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push([args[i], scalar(args[i + 1], ctx)]);
      const mask = criteriaMask(pairs, ctx, args[0]);
      if (isError(mask)) return mask;
      const vals = rangeList(args[0], ctx);
      let t = 0;
      mask.forEach((m, i) => { if (m && typeof vals[i] === 'number') t += vals[i]; });
      return round15(t);
    },
    AVERAGEIF: (args, ctx) => {
      const avgRange = args[2] || args[0];
      const mask = criteriaMask([[args[0], scalar(args[1], ctx)]], ctx);
      if (isError(mask)) return mask;
      const vals = rangeList(avgRange, ctx);
      const pick = vals.filter((v, i) => mask[i] && typeof v === 'number');
      return pick.length ? round15(pick.reduce((a, b) => a + b, 0) / pick.length) : err(ERR.DIV0, 'No cells matched the criteria');
    },
    AVERAGEIFS: (args, ctx) => {
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push([args[i], scalar(args[i + 1], ctx)]);
      const mask = criteriaMask(pairs, ctx, args[0]);
      if (isError(mask)) return mask;
      const vals = rangeList(args[0], ctx);
      const pick = vals.filter((v, i) => mask[i] && typeof v === 'number');
      return pick.length ? round15(pick.reduce((a, b) => a + b, 0) / pick.length) : err(ERR.DIV0, 'No cells matched the criteria');
    },
    MAXIFS: (args, ctx) => {
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push([args[i], scalar(args[i + 1], ctx)]);
      const mask = criteriaMask(pairs, ctx, args[0]);
      if (isError(mask)) return mask;
      const pick = rangeList(args[0], ctx).filter((v, i) => mask[i] && typeof v === 'number');
      return pick.length ? Math.max(...pick) : 0;
    },
    MINIFS: (args, ctx) => {
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push([args[i], scalar(args[i + 1], ctx)]);
      const mask = criteriaMask(pairs, ctx, args[0]);
      if (isError(mask)) return mask;
      const pick = rangeList(args[0], ctx).filter((v, i) => mask[i] && typeof v === 'number');
      return pick.length ? Math.min(...pick) : 0;
    },

    // ---- logic (lazy: receive ASTs)
    IF: (asts, ctx) => {
      if (asts.length < 2 || asts.length > 3) return err(ERR.VALUE, 'IF(test, value_if_true, [value_if_false])');
      const t = toBool(deref(evaluate(asts[0], ctx), ctx));
      if (isError(t)) return t;
      if (t) return deref(evaluate(asts[1], ctx), ctx) ?? 0;
      if (asts.length < 3) return false;
      const v = deref(evaluate(asts[2], ctx), ctx);
      return asts[2].type === 'missing' ? 0 : v;
    },
    IFS: (asts, ctx) => {
      for (let i = 0; i + 1 < asts.length; i += 2) {
        const t = toBool(deref(evaluate(asts[i], ctx), ctx));
        if (isError(t)) return t;
        if (t) return deref(evaluate(asts[i + 1], ctx), ctx);
      }
      return err(ERR.NA, 'No condition in IFS was TRUE');
    },
    IFERROR: (asts, ctx) => {
      const v = deref(evaluate(asts[0], ctx), ctx);
      return isError(v) ? deref(evaluate(asts[1], ctx), ctx) ?? 0 : v ?? 0;
    },
    IFNA: (asts, ctx) => {
      const v = deref(evaluate(asts[0], ctx), ctx);
      return isError(v) && v.code === ERR.NA ? deref(evaluate(asts[1], ctx), ctx) : v;
    },
    AND: (asts, ctx) => logical(asts, ctx, 'AND'),
    OR: (asts, ctx) => logical(asts, ctx, 'OR'),
    NOT: (args, ctx) => { const b = toBool(scalar(args[0], ctx)); return isError(b) ? b : !b; },
    XOR: (args, ctx) => {
      let n = 0;
      for (const { value } of iterArgs(args, ctx)) {
        if (isBlank(value) || typeof value === 'string') continue;
        const b = toBool(value); if (isError(b)) return b; if (b) n++;
      }
      return n % 2 === 1;
    },
    SWITCH: (asts, ctx) => {
      const v = deref(evaluate(asts[0], ctx), ctx);
      if (isError(v)) return v;
      let i = 1;
      for (; i + 1 < asts.length; i += 2) {
        const c = deref(evaluate(asts[i], ctx), ctx);
        if (compare(v, c) === 0 && typeof v === typeof c) return deref(evaluate(asts[i + 1], ctx), ctx);
      }
      return i < asts.length ? deref(evaluate(asts[i], ctx), ctx) : err(ERR.NA);
    },
    CHOOSE: (asts, ctx) => {
      const k = toNumber(deref(evaluate(asts[0], ctx), ctx));
      if (isError(k)) return k;
      const i = Math.trunc(k);
      if (i < 1 || i >= asts.length) return err(ERR.VALUE);
      return deref(evaluate(asts[i], ctx), ctx);
    },
    TRUE: () => true,
    FALSE: () => false,

    // ---- information
    ISBLANK: (args, ctx) => isBlank(scalar(args[0], ctx)),
    ISNUMBER: (args, ctx) => typeof scalar(args[0], ctx) === 'number',
    ISTEXT: (args, ctx) => typeof scalar(args[0], ctx) === 'string',
    ISLOGICAL: (args, ctx) => typeof scalar(args[0], ctx) === 'boolean',
    ISERROR: (args, ctx) => isError(scalar(args[0], ctx)),
    ISNA: (args, ctx) => { const v = scalar(args[0], ctx); return isError(v) && v.code === ERR.NA; },
    NA: () => err(ERR.NA),
    N: (args, ctx) => { const v = scalar(args[0], ctx); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; },
    TYPE: (args, ctx) => {
      const v = scalar(args[0], ctx);
      return typeof v === 'number' || isBlank(v) ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 16;
    },

    // ---- lookup & reference
    XLOOKUP: (args, ctx) => {
      const e = argCount(args, 3, 6); if (e) return e;
      const lookup = scalar(args[0], ctx);
      if (isError(lookup)) return lookup;
      if (!isRange(args[1]) || !isRange(args[2])) return err(ERR.VALUE, 'XLOOKUP needs ranges for lookup_array and return_array');
      const lookArr = args[1], retArr = args[2];
      const vertical = lookArr.cols === 1;
      const list = rangeList(lookArr, ctx);
      if ((vertical && retArr.rows !== lookArr.rows) || (!vertical && retArr.cols !== lookArr.cols)) {
        return err(ERR.VALUE, 'lookup_array and return_array must be the same length');
      }
      const matchMode = args[4] === undefined || isBlank(args[4]) ? 0 : num(args[4], ctx);
      const searchMode = args[5] === undefined || isBlank(args[5]) ? 1 : num(args[5], ctx);
      const idx = xmatchIndex(lookup, list, matchMode, searchMode);
      if (idx < 0) {
        if (args[3] !== undefined && !(args[3] === null)) return scalar(args[3], ctx);
        return err(ERR.NA, 'No match for "' + toText(lookup) + '"');
      }
      if (vertical) {
        if (retArr.cols === 1) return ctx.getCell(retArr.sheet, retArr.r1 + idx, retArr.c1);
        return ctx.getCell(retArr.sheet, retArr.r1 + idx, retArr.c1); // first column (spill not supported)
      }
      return ctx.getCell(retArr.sheet, retArr.r1, retArr.c1 + idx);
    },
    XMATCH: (args, ctx) => {
      const lookup = scalar(args[0], ctx);
      if (!isRange(args[1])) return err(ERR.VALUE);
      const matchMode = args[2] === undefined ? 0 : num(args[2], ctx);
      const searchMode = args[3] === undefined ? 1 : num(args[3], ctx);
      const idx = xmatchIndex(lookup, rangeList(args[1], ctx), matchMode, searchMode);
      return idx < 0 ? err(ERR.NA) : idx + 1;
    },
    VLOOKUP: (args, ctx) => {
      const e = argCount(args, 3, 4); if (e) return e;
      const lookup = scalar(args[0], ctx);
      if (isError(lookup)) return lookup;
      const table = args[1];
      if (!isRange(table)) return err(ERR.VALUE);
      const colIdx = num(args[2], ctx);
      if (isError(colIdx)) return colIdx;
      if (colIdx < 1) return err(ERR.VALUE);
      if (colIdx > table.cols) return err(ERR.REF, 'col_index_num is wider than the table');
      const approx = args[3] === undefined || isBlank(args[3]) ? true : toBool(scalar(args[3], ctx));
      const firstCol = [];
      for (let r = table.r1; r <= table.r2; r++) firstCol.push(ctx.getCell(table.sheet, r, table.c1));
      const idx = approx ? approxMatchAscending(firstCol, lookup) : firstCol.findIndex((v) => lookupEquals(v, lookup, true));
      if (idx < 0) return err(ERR.NA, 'No match for "' + toText(lookup) + '"');
      return ctx.getCell(table.sheet, table.r1 + idx, table.c1 + Math.trunc(colIdx) - 1);
    },
    HLOOKUP: (args, ctx) => {
      const lookup = scalar(args[0], ctx);
      const table = args[1];
      if (!isRange(table)) return err(ERR.VALUE);
      const rowIdx = num(args[2], ctx);
      if (isError(rowIdx)) return rowIdx;
      if (rowIdx > table.rows) return err(ERR.REF);
      const approx = args[3] === undefined ? true : toBool(scalar(args[3], ctx));
      const firstRow = [];
      for (let c = table.c1; c <= table.c2; c++) firstRow.push(ctx.getCell(table.sheet, table.r1, c));
      const idx = approx ? approxMatchAscending(firstRow, lookup) : firstRow.findIndex((v) => lookupEquals(v, lookup, true));
      if (idx < 0) return err(ERR.NA);
      return ctx.getCell(table.sheet, table.r1 + Math.trunc(rowIdx) - 1, table.c1 + idx);
    },
    INDEX: (args, ctx) => {
      const e = argCount(args, 2, 3); if (e) return e;
      const r = args[0];
      if (!isRange(r)) return err(ERR.VALUE);
      let row = num(args[1], ctx);
      let col = args[2] === undefined ? 1 : num(args[2], ctx);
      if (firstError(row, col)) return firstError(row, col);
      if (r.rows === 1 && args[2] === undefined) { col = row; row = 1; }
      if (row < 1 || col < 1 || row > r.rows || col > r.cols) return err(ERR.REF, 'INDEX position is outside the range');
      return ctx.getCell(r.sheet, r.r1 + Math.trunc(row) - 1, r.c1 + Math.trunc(col) - 1);
    },
    MATCH: (args, ctx) => {
      const e = argCount(args, 2, 3); if (e) return e;
      const lookup = scalar(args[0], ctx);
      if (isError(lookup)) return lookup;
      if (!isRange(args[1])) return err(ERR.VALUE);
      const list = rangeList(args[1], ctx);
      const type = args[2] === undefined ? 1 : num(args[2], ctx);
      let idx;
      if (type === 0) idx = list.findIndex((v) => lookupEquals(v, lookup, true));
      else if (type === 1) idx = approxMatchAscending(list, lookup);
      else {
        idx = -1;
        for (let i = 0; i < list.length; i++) {
          if (typeof list[i] !== typeof lookup) continue;
          if (compare(list[i], lookup) >= 0) idx = i; else break;
        }
      }
      return idx < 0 ? err(ERR.NA, 'No match for "' + toText(lookup) + '"') : idx + 1;
    },
    ROW: (args, ctx) => (args.length && isRange(args[0]) ? args[0].r1 + 1 : ctx.row + 1),
    COLUMN: (args, ctx) => (args.length && isRange(args[0]) ? args[0].c1 + 1 : ctx.col + 1),
    ROWS: (args) => (isRange(args[0]) ? args[0].rows : 1),
    COLUMNS: (args) => (isRange(args[0]) ? args[0].cols : 1),

    // ---- text
    LEN: (args, ctx) => { const s = txt(args[0], ctx); return isError(s) ? s : s.length; },
    TRIM: (args, ctx) => { const s = txt(args[0], ctx); return isError(s) ? s : s.replace(/ +/g, ' ').trim(); },
    CLEAN: (args, ctx) => { const s = txt(args[0], ctx); return isError(s) ? s : s.replace(/[\x00-\x1f]/g, ''); },
    UPPER: (args, ctx) => { const s = txt(args[0], ctx); return isError(s) ? s : s.toUpperCase(); },
    LOWER: (args, ctx) => { const s = txt(args[0], ctx); return isError(s) ? s : s.toLowerCase(); },
    PROPER: (args, ctx) => {
      const s = txt(args[0], ctx);
      if (isError(s)) return s;
      return s.toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (m, p, ch) => p + ch.toUpperCase());
    },
    LEFT: (args, ctx) => {
      const s = txt(args[0], ctx), n = args[1] === undefined ? 1 : num(args[1], ctx);
      if (firstError(s, n)) return firstError(s, n);
      return n < 0 ? err(ERR.VALUE) : s.slice(0, Math.trunc(n));
    },
    RIGHT: (args, ctx) => {
      const s = txt(args[0], ctx), n = args[1] === undefined ? 1 : num(args[1], ctx);
      if (firstError(s, n)) return firstError(s, n);
      if (n < 0) return err(ERR.VALUE);
      return Math.trunc(n) === 0 ? '' : s.slice(-Math.trunc(n));
    },
    MID: (args, ctx) => {
      const s = txt(args[0], ctx), start = num(args[1], ctx), n = num(args[2], ctx);
      if (firstError(s, start, n)) return firstError(s, start, n);
      if (start < 1 || n < 0) return err(ERR.VALUE);
      return s.substr(Math.trunc(start) - 1, Math.trunc(n));
    },
    FIND: (args, ctx) => {
      const needle = txt(args[0], ctx), hay = txt(args[1], ctx);
      const start = args[2] === undefined ? 1 : num(args[2], ctx);
      if (firstError(needle, hay, start)) return firstError(needle, hay, start);
      const i = hay.indexOf(needle, Math.trunc(start) - 1);
      return i < 0 ? err(ERR.VALUE, '"' + needle + '" was not found') : i + 1;
    },
    SEARCH: (args, ctx) => {
      const needle = txt(args[0], ctx), hay = txt(args[1], ctx);
      const start = args[2] === undefined ? 1 : num(args[2], ctx);
      if (firstError(needle, hay, start)) return firstError(needle, hay, start);
      const i = hay.toLowerCase().indexOf(needle.toLowerCase(), Math.trunc(start) - 1);
      return i < 0 ? err(ERR.VALUE, '"' + needle + '" was not found') : i + 1;
    },
    SUBSTITUTE: (args, ctx) => {
      const s = txt(args[0], ctx), from = txt(args[1], ctx), to = txt(args[2], ctx);
      if (firstError(s, from, to)) return firstError(s, from, to);
      if (from === '') return s;
      if (args[3] === undefined) return s.split(from).join(to);
      const nth = num(args[3], ctx);
      let count = 0;
      return s.replace(new RegExp(escapeRe(from), 'g'), (m) => (++count === nth ? to : m));
    },
    REPLACE: (args, ctx) => {
      const s = txt(args[0], ctx), start = num(args[1], ctx), n = num(args[2], ctx), rep = txt(args[3], ctx);
      if (firstError(s, start, n, rep)) return firstError(s, start, n, rep);
      return s.slice(0, start - 1) + rep + s.slice(start - 1 + n);
    },
    REPT: (args, ctx) => {
      const s = txt(args[0], ctx), n = num(args[1], ctx);
      return firstError(s, n) || s.repeat(Math.max(0, Math.trunc(n)));
    },
    EXACT: (args, ctx) => {
      const a = txt(args[0], ctx), b = txt(args[1], ctx);
      return firstError(a, b) || a === b;
    },
    CONCATENATE: (args, ctx) => {
      let out = '';
      for (const a of args) { const s = txt(a, ctx); if (isError(s)) return s; out += s; }
      return out;
    },
    CONCAT: (args, ctx) => {
      let out = '';
      for (const { value } of iterArgs(args, ctx)) { if (isError(value)) return value; out += toText(value); }
      return out;
    },
    TEXTJOIN: (args, ctx) => {
      const delim = txt(args[0], ctx);
      const ignore = toBool(scalar(args[1], ctx));
      if (firstError(delim, ignore)) return firstError(delim, ignore);
      const parts = [];
      for (const { value } of iterArgs(args.slice(2), ctx)) {
        if (isError(value)) return value;
        const s = toText(value);
        if (ignore && s === '') continue;
        parts.push(s);
      }
      return parts.join(delim);
    },
    TEXT: (args, ctx) => {
      const v = scalar(args[0], ctx), code = txt(args[1], ctx);
      if (firstError(v, code)) return firstError(v, code);
      if (typeof v === 'string' && !parseNumberLike(v)) return v;
      const n = toNumber(v);
      if (isError(n)) return n;
      return applyFormatCode(n, code);
    },
    VALUE: (args, ctx) => {
      const v = scalar(args[0], ctx);
      if (typeof v === 'number') return v;
      const p = parseNumberLike(toText(v));
      return p ? p.value : err(ERR.VALUE, '"' + toText(v) + '" cannot be turned into a number');
    },
    NUMBERVALUE: (args, ctx) => FUNCTIONS.VALUE(args, ctx),
    CHAR: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : String.fromCharCode(n); },
    CODE: (args, ctx) => { const s = txt(args[0], ctx); return isError(s) ? s : s.length ? s.charCodeAt(0) : err(ERR.VALUE); },

    // ---- dates
    DATE: (args, ctx) => {
      const y = num(args[0], ctx), m = num(args[1], ctx), d = num(args[2], ctx);
      if (firstError(y, m, d)) return firstError(y, m, d);
      const year = y < 1900 ? y + 1900 : y;
      return dateToSerial(Math.trunc(year), Math.trunc(m), Math.trunc(d));
    },
    YEAR: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : serialToDate(n).y; },
    MONTH: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : serialToDate(n).m; },
    DAY: (args, ctx) => { const n = num(args[0], ctx); return isError(n) ? n : serialToDate(n).d; },
    WEEKDAY: (args, ctx) => {
      const n = num(args[0], ctx);
      if (isError(n)) return n;
      const type = args[1] === undefined ? 1 : num(args[1], ctx);
      const dow = serialToDate(n).dow; // 0 = Sunday
      if (type === 2) return dow === 0 ? 7 : dow;
      if (type === 3) return dow === 0 ? 6 : dow - 1;
      return dow + 1;
    },
    TODAY: () => todaySerial(),
    DATEDIF: (args, ctx) => {
      const a = num(args[0], ctx), b = num(args[1], ctx), unit = txt(args[2], ctx);
      if (firstError(a, b, unit)) return firstError(a, b, unit);
      if (a > b) return err(ERR.NUM, 'Start date is after end date');
      const s = serialToDate(a), e = serialToDate(b);
      switch (unit.toUpperCase()) {
        case 'D': return Math.floor(b) - Math.floor(a);
        case 'M': return (e.y - s.y) * 12 + (e.m - s.m) - (e.d < s.d ? 1 : 0);
        case 'Y': return e.y - s.y - (e.m < s.m || (e.m === s.m && e.d < s.d) ? 1 : 0);
        default: return err(ERR.NUM, 'Unit must be "Y", "M" or "D"');
      }
    },
    EDATE: (args, ctx) => {
      const a = num(args[0], ctx), m = num(args[1], ctx);
      if (firstError(a, m)) return firstError(a, m);
      const s = serialToDate(a);
      const target = dateToSerial(s.y, s.m + Math.trunc(m), 1);
      const t = serialToDate(target);
      const last = serialToDate(dateToSerial(t.y, t.m + 1, 0)).d;
      return dateToSerial(t.y, t.m, Math.min(s.d, last));
    },
    EOMONTH: (args, ctx) => {
      const a = num(args[0], ctx), m = num(args[1], ctx);
      if (firstError(a, m)) return firstError(a, m);
      const s = serialToDate(a);
      return dateToSerial(s.y, s.m + Math.trunc(m) + 1, 0);
    },
    DAYS: (args, ctx) => {
      const e = num(args[0], ctx), s = num(args[1], ctx);
      return firstError(e, s) || Math.floor(e) - Math.floor(s);
    },
  };

  function logical(asts, ctx, mode) {
    let seen = false;
    for (const a of asts) {
      const v = evaluate(a, ctx);
      const vals = isRange(v) ? rangeList(v, ctx) : [v];
      for (const x of vals) {
        if (isError(x)) return x;
        if (isBlank(x) || (isRange(v) && typeof x === 'string')) continue;
        const b = toBool(x);
        if (isError(b)) return b;
        seen = true;
        if (mode === 'AND' && !b) return false;
        if (mode === 'OR' && b) return true;
      }
    }
    if (!seen) return err(ERR.VALUE);
    return mode === 'AND';
  }

  const SUPPORTED_FUNCTIONS = Object.keys(FUNCTIONS).sort();

  const api = {
    XLError, RangeRef, ERR, err, isError, isBlank, isRange,
    colToIndex, indexToCol, toAddress, parseAddress, parseRangeAddress, rangeToAddress, quoteSheet,
    dateToSerial, serialToDate, todaySerial, parseNumberLike, round15,
    toNumber, toText, toBool, compare,
    formatGeneral, formatValue, applyFormatCode, NAMED_FORMATS,
    tokenize, parse, evaluate, deref, functionsUsed, shiftFormula, referencedRanges,
    makeCriterion, SUPPORTED_FUNCTIONS,
  };

  root.DIM = root.DIM || {};
  root.DIM.Formula = api;
  if (typeof module === 'object' && module.exports) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
