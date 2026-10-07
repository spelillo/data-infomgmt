/*
 * xlsx.js — read and write real Excel (.xlsx) files with no dependencies.
 *
 * Writing: builds the Office Open XML parts (cells, number formats, column
 * widths, defined names, data validation, conditional formatting and Excel
 * Tables) and stores them in a ZIP container
 * (uncompressed "stored" entries, which every spreadsheet app accepts).
 * Formulas are written with their cached values and the workbook is flagged to
 * fully recalculate on open, so Excel, Google Sheets and LibreOffice all show
 * live results.
 *
 * Reading: unzips (inflate via zlib in Node, DecompressionStream in browsers),
 * then parses sharedStrings, styles (to recover number formats), defined names
 * and each worksheet, including Excel's "shared formulas" created by fill-down.
 * It also reads the structures students build in Excel so uploads can be
 * graded: Tables, data validation, conditional formatting, charts,
 * PivotTables (with date grouping) and Power Query queries (decoded from the
 * DataMashup package and mapped from M onto this site's Applied Steps).
 */
(function (root) {
  'use strict';

  const isNode = typeof module === 'object' && module.exports;
  const F = isNode ? require('./formula.js') : root.DIM.Formula;
  const W = isNode ? require('./workbook.js') : root.DIM.Workbook;

  // ---------------------------------------------------------------------------
  // CRC32 + ZIP
  // ---------------------------------------------------------------------------

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  const enc = new TextEncoder();
  const dec = new TextDecoder('utf-8');

  function zipStore(files) {
    // files: [{ name, data: Uint8Array }]
    const chunks = [];
    const central = [];
    let offset = 0;
    const dosTime = 0, dosDate = (2026 - 1980) << 9 | 1 << 5 | 1;
    for (const f of files) {
      const name = enc.encode(f.name);
      const crc = crc32(f.data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true);
      local.setUint16(4, 20, true);
      local.setUint16(6, 0x0800, true); // UTF-8 names
      local.setUint16(8, 0, true); // stored
      local.setUint16(10, dosTime, true);
      local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true);
      local.setUint32(18, f.data.length, true);
      local.setUint32(22, f.data.length, true);
      local.setUint16(26, name.length, true);
      local.setUint16(28, 0, true);
      chunks.push(new Uint8Array(local.buffer), name, f.data);
      const cen = new DataView(new ArrayBuffer(46));
      cen.setUint32(0, 0x02014b50, true);
      cen.setUint16(4, 20, true);
      cen.setUint16(6, 20, true);
      cen.setUint16(8, 0x0800, true);
      cen.setUint16(10, 0, true);
      cen.setUint16(12, dosTime, true);
      cen.setUint16(14, dosDate, true);
      cen.setUint32(16, crc, true);
      cen.setUint32(20, f.data.length, true);
      cen.setUint32(24, f.data.length, true);
      cen.setUint16(28, name.length, true);
      cen.setUint32(42, offset, true);
      central.push(new Uint8Array(cen.buffer), name);
      offset += 30 + name.length + f.data.length;
    }
    const cenSize = central.reduce((n, c) => n + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cenSize, true);
    end.setUint32(16, offset, true);
    const all = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((n, c) => n + c.length, 0));
    let p = 0;
    for (const c of all) { out.set(c, p); p += c.length; }
    return out;
  }

  async function inflateRaw(data) {
    if (isNode) return new Uint8Array(require('zlib').inflateRawSync(data));
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([data]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  async function unzip(bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This file is not a valid .xlsx (ZIP) file.');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Corrupt ZIP directory.');
      const method = dv.getUint16(p + 10, true);
      const compSize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const localOff = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nameLen));
      const lNameLen = dv.getUint16(localOff + 26, true);
      const lExtraLen = dv.getUint16(localOff + 28, true);
      const start = localOff + 30 + lNameLen + lExtraLen;
      const raw = u8.subarray(start, start + compSize);
      files[name] = { method, raw };
      p += 46 + nameLen + extraLen + commentLen;
    }
    const out = {};
    for (const [name, f] of Object.entries(files)) {
      if (f.method === 0) out[name] = f.raw;
      else if (f.method === 8) out[name] = await inflateRaw(f.raw);
      else throw new Error('Unsupported ZIP compression in ' + name);
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // XML helpers
  // ---------------------------------------------------------------------------

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  }

  function unesc(s) {
    return String(s).replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, e) => {
      if (e === 'lt') return '<';
      if (e === 'gt') return '>';
      if (e === 'amp') return '&';
      if (e === 'quot') return '"';
      if (e === 'apos') return "'";
      if (e[1] === 'x' || e[1] === 'X') return String.fromCodePoint(parseInt(e.slice(2), 16));
      return String.fromCodePoint(parseInt(e.slice(1), 10));
    }).replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  }

  function attrs(tag) {
    const o = {};
    const re = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let m;
    while ((m = re.exec(tag))) o[m[1].replace(/^\w+:/, (p) => (p === 'r:' ? 'r:' : ''))] = unesc(m[3] !== undefined ? m[3] : m[4]);
    return o;
  }

  // Concatenate the text of every <t> inside a fragment (handles rich text runs).
  function textOfT(fragment) {
    let s = '';
    const re = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>|<(?:\w+:)?t(?:\s[^>]*)?\/>/g;
    let m;
    while ((m = re.exec(fragment))) s += m[1] !== undefined ? unesc(m[1]) : '';
    return s;
  }

  // Strip <rPh> phonetic runs which should not appear in the text.
  function stripPhonetic(fragment) { return fragment.replace(/<(?:\w+:)?rPh[\s\S]*?<\/(?:\w+:)?rPh>/g, ''); }

  // ---------------------------------------------------------------------------
  // Styles: map our named formats <-> Excel number formats
  // ---------------------------------------------------------------------------

  const FORMAT_ORDER = ['general', 'integer', 'number', 'currency', 'percent', 'percent1', 'date', 'longdate', 'text'];
  const BUILTIN_IDS = { general: 0, integer: 3, number: 4, percent: 9, date: 14, text: 49 };
  const CUSTOM = { currency: { id: 164, code: '"$"#,##0.00' }, percent1: { id: 165, code: '0.0%' }, longdate: { id: 166, code: 'mmmm d, yyyy' } };

  // Conditional-format looks, matching Excel's built-in presets.
  const DXF = { green: { fill: 'FFC6EFCE', font: 'FF006100' }, red: { fill: 'FFFFC7CE', font: 'FF9C0006' }, yellow: { fill: 'FFFFEB9C', font: 'FF9C5700' } };
  const DXF_ORDER = ['green', 'red', 'yellow'];

  function styleIndex(format, bold) {
    const i = Math.max(0, FORMAT_ORDER.indexOf(format || 'general'));
    return i * 2 + (bold ? 1 : 0);
  }

  function stylesXml() {
    const numFmts = Object.values(CUSTOM).map((c) => '<numFmt numFmtId="' + c.id + '" formatCode="' + esc(c.code) + '"/>').join('');
    const xfs = [];
    for (const f of FORMAT_ORDER) {
      const id = BUILTIN_IDS[f] !== undefined ? BUILTIN_IDS[f] : CUSTOM[f].id;
      for (const bold of [0, 1]) {
        xfs.push('<xf numFmtId="' + id + '" fontId="' + bold + '" fillId="0" borderId="0" xfId="0"' + (id ? ' applyNumberFormat="1"' : '') + (bold ? ' applyFont="1"' : '') + '/>');
      }
    }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="' + Object.keys(CUSTOM).length + '">' + numFmts + '</numFmts>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="' + DXF_ORDER.length + '">' + DXF_ORDER.map((k) => '<dxf><font><color rgb="' + DXF[k].font + '"/></font><fill><patternFill><bgColor rgb="' + DXF[k].fill + '"/></patternFill></fill></dxf>').join('') + '</dxfs>' +
      '</styleSheet>';
  }

  // Interpret an Excel number format id/code as one of our named formats.
  function namedFormatFor(id, code) {
    id = Number(id);
    if (id === 0) return null;
    if (id === 49 || code === '@') return 'text';
    if (id === 9) return 'percent';
    if (id === 10) return 'percent1';
    if (id === 3 || id === 37 || id === 38) return 'integer';
    if (id === 4 || id === 39 || id === 40) return 'number';
    if (id === 5 || id === 6 || id === 7 || id === 8 || id === 44) return 'currency';
    if ((id >= 14 && id <= 17) || id === 22) return 'date';
    if (!code) return null;
    const c = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
    if (/%/.test(c)) return /0\.0/.test(c) ? 'percent1' : 'percent';
    if (/[$€£]/.test(code)) return 'currency';
    if (/[dy]/i.test(c) || /mmm/i.test(c)) return /mmmm/i.test(c) ? 'longdate' : 'date';
    if (/#,##0\.0/.test(c)) return 'number';
    if (/#,##0/.test(c)) return 'integer';
    if (/0\.0/.test(c)) return 'number';
    return null;
  }

  // ---------------------------------------------------------------------------
  // Writing
  // ---------------------------------------------------------------------------

  function cellXml(wb, sheet, key, cell, keepTables) {
    const a = F.parseAddress(key);
    const fmt = cell.format || (W.interpretInput(cell.input).autoFormat) || 'general';
    const s = styleIndex(fmt, cell.style && cell.style.bold);
    const sAttr = s ? ' s="' + s + '"' : '';
    const input = cell.input;
    if (input === '' || input === null || input === undefined) return s ? '<c r="' + key + '"' + sAttr + '/>' : '';
    const value = wb.getValue(sheet.name, a);
    if (typeof input === 'string' && input.startsWith('=') && input.length > 1) {
      const formula = wb.formulaToA1(sheet.name, a.row, a.col, input, { keepTables: keepTables || null }).slice(1);
      let v = '', t = '';
      if (typeof value === 'number') v = '<v>' + value + '</v>';
      else if (typeof value === 'boolean') { t = ' t="b"'; v = '<v>' + (value ? 1 : 0) + '</v>'; }
      else if (F.isError(value)) { t = ' t="e"'; v = '<v>' + esc(value.code === '#CIRC!' ? '#VALUE!' : value.code) + '</v>'; }
      else if (typeof value === 'string') { t = ' t="str"'; v = '<v>' + esc(value) + '</v>'; }
      return '<c r="' + key + '"' + sAttr + t + '><f>' + esc(formula) + '</f>' + v + '</c>';
    }
    if (typeof value === 'number') return '<c r="' + key + '"' + sAttr + '><v>' + value + '</v></c>';
    if (typeof value === 'boolean') return '<c r="' + key + '"' + sAttr + ' t="b"><v>' + (value ? 1 : 0) + '</v></c>';
    if (F.isError(value)) return '<c r="' + key + '"' + sAttr + ' t="e"><v>' + esc(value.code) + '</v></c>';
    const str = String(value ?? '');
    const sp = /^\s|\s$|\n/.test(str) ? ' xml:space="preserve"' : '';
    return '<c r="' + key + '"' + sAttr + ' t="inlineStr"><is><t' + sp + '>' + esc(str) + '</t></is></c>';
  }

  function sheetXml(wb, sheet, tableRids, keepTables) {
    const rows = new Map();
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const a = F.parseAddress(key);
      if (!a) continue;
      const xml = cellXml(wb, sheet, key, cell, keepTables);
      if (!xml) continue;
      if (!rows.has(a.row)) rows.set(a.row, []);
      rows.get(a.row).push({ col: a.col, xml });
    }
    const sortedRows = [...rows.entries()].sort((x, y) => x[0] - y[0]);
    const body = sortedRows.map(([r, cells]) => '<row r="' + (r + 1) + '">' + cells.sort((x, y) => x.col - y.col).map((c) => c.xml).join('') + '</row>').join('');
    const widths = Object.entries(sheet.widths || {}).map(([col, px]) => {
      const i = F.colToIndex(col) + 1;
      return '<col min="' + i + '" max="' + i + '" width="' + Math.max(4, Math.round((px / 7) * 100) / 100) + '" customWidth="1"/>';
    }).join('');
    const b = wb.usedBounds(sheet.name);
    const dim = b.rows && b.cols ? 'A1:' + F.toAddress(b.rows - 1, b.cols - 1) : 'A1';
    let validations = '';
    if (sheet.validations && sheet.validations.length) {
      const items = sheet.validations.map((v) => {
        if (v.type === 'list') {
          return '<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="' + esc(v.range) + '"><formula1>"' + esc(v.values.join(',')) + '"</formula1></dataValidation>';
        }
        if (v.type === 'whole' || v.type === 'decimal') {
          return '<dataValidation type="' + v.type + '" operator="between" allowBlank="1" showErrorMessage="1" sqref="' + esc(v.range) + '"><formula1>' + (v.min ?? -1e9) + '</formula1><formula2>' + (v.max ?? 1e9) + '</formula2></dataValidation>';
        }
        return '';
      }).filter(Boolean);
      if (items.length) validations = '<dataValidations count="' + items.length + '">' + items.join('') + '</dataValidations>';
    }
    const condFormats = condFormatXml(sheet);
    const tableParts = tableRids && tableRids.length
      ? '<tableParts count="' + tableRids.length + '">' + tableRids.map((id) => '<tablePart r:id="' + id + '"/>').join('') + '</tableParts>' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<dimension ref="' + dim + '"/>' +
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      (widths ? '<cols>' + widths + '</cols>' : '') +
      '<sheetData>' + body + '</sheetData>' + condFormats + validations + tableParts +
      '</worksheet>';
  }

  const CF_XML_OPS = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '<>': 'notEqual', between: 'between' };

  function cfFormula(v) {
    return typeof v === 'number' || (v !== '' && v !== null && v !== undefined && !Number.isNaN(Number(v))) ? String(Number(v)) : '"' + String(v ?? '').replace(/"/g, '""') + '"';
  }

  function condFormatXml(sheet) {
    let priority = 0;
    return (sheet.condFormats || []).map((cf) => {
      const r = F.parseRangeAddress(cf.range);
      if (!r) return '';
      const tl = F.toAddress(r.r1, r.c1);
      const dxf = Math.max(0, DXF_ORDER.indexOf(cf.style || 'red'));
      const head = '<cfRule dxfId="' + dxf + '" priority="' + (++priority) + '" ';
      let rule;
      if (CF_XML_OPS[cf.op]) {
        rule = head + 'type="cellIs" operator="' + CF_XML_OPS[cf.op] + '"><formula>' + esc(cfFormula(cf.value)) + '</formula>' +
          (cf.op === 'between' ? '<formula>' + esc(cfFormula(cf.value2)) + '</formula>' : '') + '</cfRule>';
      } else if (cf.op === 'contains') {
        rule = head + 'type="containsText" operator="containsText" text="' + esc(cf.value) + '"><formula>' + esc('NOT(ISERROR(SEARCH(' + cfFormula(String(cf.value)) + ',' + tl + ')))') + '</formula></cfRule>';
      } else if (cf.op === 'duplicate') rule = head + 'type="duplicateValues"/>';
      else if (cf.op === 'top') rule = head + 'type="top10" rank="' + (Number(cf.value) || 10) + '"/>';
      else if (cf.op === 'aboveAverage') rule = head + 'type="aboveAverage"/>';
      else if (cf.op === 'blank') rule = head + 'type="containsBlanks"><formula>' + esc('LEN(TRIM(' + tl + '))=0') + '</formula></cfRule>';
      else return '';
      return '<conditionalFormatting sqref="' + esc(F.rangeToAddress(r)) + '">' + rule + '</conditionalFormatting>';
    }).join('');
  }

  // A table can be written only when its header row is unique, non-empty text
  // (Excel requires the header cells to match the table's column names).
  function tableColumns(wb, t) {
    const r = F.parseRangeAddress(t.ref);
    if (!r || r.r2 <= r.r1) return null;
    const names = [];
    for (let c = r.c1; c <= r.c2; c++) {
      const v = wb.getValue(t.sheet, { row: r.r1, col: c });
      if (typeof v !== 'string' || !v.trim()) return null;
      if (names.some((n) => n.toLowerCase() === v.toLowerCase())) return null;
      names.push(v);
    }
    return names;
  }

  function tableXml(t, id, columns) {
    const ref = F.rangeToAddress(F.parseRangeAddress(t.ref));
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="' + id + '" name="' + esc(t.name) + '" displayName="' + esc(t.name) + '" ref="' + ref + '" totalsRowShown="0">' +
      '<autoFilter ref="' + ref + '"/><tableColumns count="' + columns.length + '">' +
      columns.map((c, i) => '<tableColumn id="' + (i + 1) + '" name="' + esc(c) + '"/>').join('') +
      '</tableColumns><tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
  }

  // Returns a Uint8Array containing a valid .xlsx file.
  function write(wb, opts = {}) {
    const sheets = wb.sheets.filter((s) => !(opts.exclude || []).includes(s.name));
    const files = [];
    const add = (name, text) => files.push({ name, data: enc.encode(text) });
    // tables per sheet
    const tablesBySheet = sheets.map(() => []);
    let tableId = 0;
    for (const t of wb.tables || []) {
      const si = sheets.findIndex((s) => s.name === t.sheet);
      const columns = si >= 0 ? tableColumns(wb, t) : null;
      if (columns) tablesBySheet[si].push({ t, id: ++tableId, columns });
    }
    const keepTables = new Set(tablesBySheet.flat().map((x) => x.t.name));
    add('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') +
      tablesBySheet.flat().map((x) => '<Override PartName="/xl/tables/table' + x.id + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>').join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      '</Types>');
    add('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>');
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    add('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:title>' + esc(opts.title || 'Workbook') + '</dc:title><dc:creator>Intro to Data and Information Management</dc:creator>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>');
    add('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
    const names = Object.entries(wb.names || {}).map(([n, ref]) => {
      const r = F.parseRangeAddress(ref);
      const sheet = r && r.sheet ? r.sheet : sheets[0].name;
      const a = F.toAddress(r.r1, r.c1, true, true), b = F.toAddress(r.r2, r.c2, true, true);
      return '<definedName name="' + esc(n) + '">' + esc(F.quoteSheet(sheet) + '!' + (a === b ? a : a + ':' + b)) + '</definedName>';
    }).join('');
    add('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView/></bookViews><sheets>' +
      sheets.map((s, i) => '<sheet name="' + esc(s.name.slice(0, 31)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') +
      '</sheets>' + (names ? '<definedNames>' + names + '</definedNames>' : '') +
      '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    add('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>');
    add('xl/styles.xml', stylesXml());
    sheets.forEach((s, i) => {
      const ts = tablesBySheet[i];
      add('xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(wb, s, ts.map((x, j) => 'rId' + (j + 1)), keepTables));
      if (!ts.length) return;
      add('xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        ts.map((x, j) => '<Relationship Id="rId' + (j + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table' + x.id + '.xml"/>').join('') +
        '</Relationships>');
      ts.forEach((x) => add('xl/tables/table' + x.id + '.xml', tableXml(x.t, x.id, x.columns)));
    });
    return zipStore(files);
  }

  // ---------------------------------------------------------------------------
  // Reading
  // ---------------------------------------------------------------------------

  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/').slice(0, -1);
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  }

  function readText(files, path) {
    const f = files[path] || files[Object.keys(files).find((k) => k.toLowerCase() === path.toLowerCase())];
    return f ? dec.decode(f) : null;
  }

  // Parse an .xlsx (ArrayBuffer/Uint8Array) into a Workbook.
  async function read(data) {
    const files = await unzip(data);
    const wbXml = readText(files, 'xl/workbook.xml');
    if (!wbXml) throw new Error('This file does not look like an Excel workbook (xl/workbook.xml is missing).');
    const relsXml = readText(files, 'xl/_rels/workbook.xml.rels') || '';
    const rels = {};
    (relsXml.match(/<Relationship\b[^>]*>/g) || []).forEach((tag) => { const a = attrs(tag); rels[a.Id] = a.Target; });

    // shared strings
    const sst = [];
    const sstXml = readText(files, 'xl/sharedStrings.xml');
    if (sstXml) {
      const re = /<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>|<(?:\w+:)?si\/>/g;
      let m;
      while ((m = re.exec(sstXml))) sst.push(m[1] !== undefined ? textOfT(stripPhonetic(m[1])) : '');
    }

    // styles -> named format per xf index
    const xfFormats = [];
    const xfBold = [];
    const stylesXmlText = readText(files, 'xl/styles.xml');
    if (stylesXmlText) {
      const codes = {};
      (stylesXmlText.match(/<(?:\w+:)?numFmt\b[^>]*>/g) || []).forEach((tag) => { const a = attrs(tag); codes[a.numFmtId] = a.formatCode; });
      const boldFonts = [];
      const fontsBlock = /<(?:\w+:)?fonts\b[^>]*>([\s\S]*?)<\/(?:\w+:)?fonts>/.exec(stylesXmlText);
      if (fontsBlock) {
        const fre = /<(?:\w+:)?font\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?font>)/g;
        let fm;
        while ((fm = fre.exec(fontsBlock[1]))) boldFonts.push(/<(?:\w+:)?b(?:\s|\/|>)/.test(fm[1] || '') && !/<(?:\w+:)?b\s+val="(0|false)"/.test(fm[1] || ''));
      }
      const cellXfs = /<(?:\w+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/.exec(stylesXmlText);
      if (cellXfs) {
        (cellXfs[1].match(/<(?:\w+:)?xf\b[^>]*>/g) || []).forEach((tag) => {
          const a = attrs(tag);
          xfFormats.push(namedFormatFor(a.numFmtId || 0, codes[a.numFmtId]));
          xfBold.push(!!boldFonts[Number(a.fontId || 0)]);
        });
      }
    }

    const wb = new W.Workbook();
    const sheetTags = wbXml.match(/<(?:\w+:)?sheet\b[^>]*>/g) || [];
    const parsed = [];
    for (const tag of sheetTags) {
      const a = attrs(tag);
      const rid = a['r:id'] || a.id;
      const target = rels[rid];
      if (!target) continue;
      const path = resolvePath('xl/workbook.xml', target);
      const xml = readText(files, path);
      if (xml === null) continue;
      const sheet = wb.addSheet(a.name, { silent: true });
      parseSheet(xml, sheet, sst, xfFormats, xfBold);
      parsed.push({ sheet, path, xml });
    }
    // defined names
    const dn = /<(?:\w+:)?definedName\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?definedName>/g;
    let m;
    while ((m = dn.exec(wbXml))) {
      const a = attrs('<x ' + m[1] + '>');
      if (a.name.startsWith('_xlnm.') || a.localSheetId !== undefined) continue;
      const ref = unesc(m[2]).replace(/\$/g, '');
      if (F.parseRangeAddress(ref)) wb.names[a.name] = unesc(m[2]);
    }
    if (!wb.sheets.length) wb.addSheet('Sheet1', { silent: true });
    wb.recalc();
    // Structures: tables first (pivots, validations and queries refer to them).
    const dxfStyles = parseDxfs(stylesXmlText || '');
    const ctx = { files, wb, pivotSheets: [] };
    for (const p of parsed) readTables(ctx, p);
    for (const p of parsed) {
      readValidations(ctx, p);
      readCondFormats(ctx, p, dxfStyles);
      readCharts(ctx, p);
      readPivots(ctx, p);
    }
    await readQueries(ctx);
    wb.recalc();
    return wb;
  }

  function parseSheet(xml, sheet, sst, xfFormats, xfBold) {
    const shared = {}; // si -> { row, col, formula }
    const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let m;
    let autoRow = 0, autoCol = -1, lastRowSeen = -1;
    const rowRe = /<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g;
    let rm;
    while ((rm = rowRe.exec(xml))) {
      const ra = attrs('<x ' + rm[1] + '>');
      const rowIndex = ra.r ? parseInt(ra.r, 10) - 1 : lastRowSeen + 1;
      lastRowSeen = rowIndex;
      autoRow = rowIndex; autoCol = -1;
      const inner = rm[2] || '';
      cellRe.lastIndex = 0;
      while ((m = cellRe.exec(inner))) {
        const a = attrs('<x ' + m[1] + '>');
        let addr;
        if (a.r) { addr = F.parseAddress(a.r); autoCol = addr.col; }
        else { autoCol++; addr = { row: autoRow, col: autoCol }; }
        const body = m[2] || '';
        const type = a.t || 'n';
        const fmt = a.s !== undefined ? xfFormats[Number(a.s)] : null;
        const bold = a.s !== undefined ? xfBold[Number(a.s)] : false;
        const fm = /<(?:\w+:)?f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?f>)/.exec(body);
        const vm = /<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(body);
        const rawV = vm ? unesc(vm[1]) : null;
        let value = null;
        if (type === 's') value = rawV !== null ? sst[Number(rawV)] ?? '' : null;
        else if (type === 'inlineStr') value = textOfT(body.replace(/<(?:\w+:)?f\b[\s\S]*?(?:\/>|<\/(?:\w+:)?f>)/, ''));
        else if (type === 'str') value = rawV ?? '';
        else if (type === 'b') value = rawV === '1' || rawV === 'true';
        else if (type === 'e') value = F.err(rawV || '#VALUE!');
        else if (rawV !== null && rawV !== '') value = Number(rawV);

        let formula = null;
        if (fm) {
          const fa = attrs('<x ' + fm[1] + '>');
          const ftext = fm[2] !== undefined ? unesc(fm[2]) : '';
          if (fa.t === 'shared') {
            if (ftext) {
              shared[fa.si] = { row: addr.row, col: addr.col, formula: '=' + ftext };
              formula = '=' + ftext;
            } else if (shared[fa.si]) {
              const base = shared[fa.si];
              formula = F.shiftFormula(base.formula, addr.row - base.row, addr.col - base.col);
            }
          } else if (fa.t === 'array' || fa.t === 'dataTable') {
            formula = ftext ? '=' + ftext : null;
          } else if (ftext) {
            formula = '=' + ftext;
          }
        }
        const key = F.toAddress(addr.row, addr.col);
        const cell = {};
        if (formula) {
          cell.input = formula.replace(/_xlfn\._xlws\.|_xlfn\./gi, '');
          if (value !== null) cell.cached = value;
        } else if (value === null) {
          if (!fmt && !bold) continue;
          cell.input = '';
        } else if (typeof value === 'string') {
          cell.input = W.inputForValue(value);
        } else if (F.isError(value)) {
          cell.input = value.code;
        } else {
          cell.input = value;
        }
        if (fmt) cell.format = fmt;
        if (bold) cell.style = { bold: true };
        sheet.cells[key] = cell;
      }
    }
    // column widths
    const colRe = /<(?:\w+:)?col\b([^>]*)\/?>/g;
    while ((m = colRe.exec(xml))) {
      const a = attrs('<x ' + m[1] + '>');
      if (!a.width || !a.min) continue;
      for (let i = Number(a.min); i <= Math.min(Number(a.max || a.min), 52); i++) sheet.widths[F.indexToCol(i - 1)] = Math.round(Number(a.width) * 7);
    }
  }

  // ---------------------------------------------------------------------------
  // Reading structures: tables, data validation, conditional formatting,
  // charts, PivotTables and Power Query. Each reader is best-effort: anything
  // it cannot understand is skipped rather than failing the whole upload, so a
  // student's file always opens and the grader reports what it could not find.
  // ---------------------------------------------------------------------------

  function relsFor(files, partPath) {
    const i = partPath.lastIndexOf('/');
    const relsPath = partPath.slice(0, i) + '/_rels/' + partPath.slice(i + 1) + '.rels';
    const xml = readText(files, relsPath);
    const out = {};
    (xml && xml.match(/<Relationship\b[^>]*>/g) || []).forEach((tag) => {
      const a = attrs(tag);
      if (a.TargetMode === 'External') return;
      out[a.Id] = { path: resolvePath(partPath, a.Target), type: String(a.Type || '').split('/').pop() };
    });
    return out;
  }

  function relsOfType(files, partPath, type) {
    return Object.values(relsFor(files, partPath)).filter((r) => r.type === type).map((r) => r.path);
  }

  // "'My Sheet'!$A$1:$B$5" -> { sheet: 'My Sheet', range: 'A1:B5' }
  function splitRef(ref) {
    const s = String(ref || '').trim().replace(/\$/g, '');
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(s);
    if (m) return { sheet: (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]), range: m[3] };
    return { sheet: null, range: s };
  }

  function readTables(ctx, p) {
    for (const path of relsOfType(ctx.files, p.path, 'table')) {
      const xml = readText(ctx.files, path);
      if (!xml) continue;
      const t = /<(?:\w+:)?table\b[^>]*>/.exec(xml);
      if (!t) continue;
      const a = attrs(t[0]);
      if (!a.ref) continue;
      const name = a.displayName || a.name || 'Table' + (ctx.wb.tables.length + 1);
      const entry = { name, sheet: p.sheet.name, ref: a.ref.replace(/\$/g, '') };
      if (a.tableType === 'queryTable') {
        // follow table -> queryTable -> connection to learn which query loaded it
        const qtPath = relsOfType(ctx.files, path, 'queryTable')[0];
        const qt = qtPath && readText(ctx.files, qtPath);
        const cm = qt && /connectionId="(\d+)"/.exec(qt);
        if (cm) entry.connectionId = cm[1];
      }
      ctx.wb.tables.push(entry);
    }
  }

  function readValidations(ctx, p) {
    const s = p.sheet;
    s.validations = s.validations || [];
    const re = /<(?:\w+:)?dataValidation\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?dataValidation>/g;
    let m;
    while ((m = re.exec(p.xml))) {
      const a = attrs('<x ' + m[1] + '>');
      const body = m[2];
      const f1m = /<(?:\w+:)?formula1>(?:\s*<xm:f>)?([\s\S]*?)(?:<\/xm:f>\s*)?<\/(?:\w+:)?formula1>/.exec(body);
      const f2m = /<(?:\w+:)?formula2>(?:\s*<xm:f>)?([\s\S]*?)(?:<\/xm:f>\s*)?<\/(?:\w+:)?formula2>/.exec(body);
      const sqm = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(body);
      const sqref = a.sqref || (sqm ? sqm[1] : '');
      const f1 = f1m ? unesc(f1m[1]).trim() : '';
      const f2 = f2m ? unesc(f2m[1]).trim() : '';
      let rule = null;
      if (a.type === 'list') {
        let values;
        if (/^".*"$/.test(f1)) values = f1.slice(1, -1).split(',').map((x) => x.trim()).filter((x) => x !== '');
        else {
          const ref = splitRef(f1);
          const named = ctx.wb.names && ctx.wb.names[ref.range];
          const target = named ? splitRef(named) : ref;
          const range = F.parseRangeAddress(target.range);
          if (range) {
            const vals = ctx.wb.getRangeValues(target.sheet || s.name, range).flat();
            values = vals.filter((v) => v !== null && v !== undefined && v !== '').map((v) => String(F.toText(v)));
          }
        }
        if (values) rule = { type: 'list', values };
      } else if (a.type === 'whole' || a.type === 'decimal') {
        const n1 = f1 !== '' && !Number.isNaN(Number(f1)) ? Number(f1) : undefined;
        const n2 = f2 !== '' && !Number.isNaN(Number(f2)) ? Number(f2) : undefined;
        const op = a.operator || 'between';
        rule = { type: a.type };
        if (op === 'between') { rule.min = n1; rule.max = n2; }
        else if (op === 'greaterThanOrEqual' || op === 'greaterThan') rule.min = n1;
        else if (op === 'lessThanOrEqual' || op === 'lessThan') rule.max = n1;
        else if (op === 'equal') { rule.min = n1; rule.max = n1; }
      }
      if (!rule) continue;
      for (const range of sqref.split(/\s+/).filter(Boolean)) s.validations.push({ range, ...rule });
    }
  }

  // Classify each differential format (dxf) by its fill as green, red or yellow.
  function parseDxfs(stylesText) {
    const block = /<(?:\w+:)?dxfs\b[^>]*>([\s\S]*?)<\/(?:\w+:)?dxfs>/.exec(stylesText);
    if (!block) return [];
    const out = [];
    const re = /<(?:\w+:)?dxf>([\s\S]*?)<\/(?:\w+:)?dxf>|<(?:\w+:)?dxf\/>/g;
    let m;
    while ((m = re.exec(block[1]))) {
      const body = m[1] || '';
      const fill = /<(?:\w+:)?fill>[\s\S]*?<(?:\w+:)?bgColor\b[^>]*rgb="([0-9A-Fa-f]{6,8})"/.exec(body)
        || /<(?:\w+:)?fill>[\s\S]*?<(?:\w+:)?fgColor\b[^>]*rgb="([0-9A-Fa-f]{6,8})"/.exec(body)
        || /<(?:\w+:)?font>[\s\S]*?<(?:\w+:)?color\b[^>]*rgb="([0-9A-Fa-f]{6,8})"/.exec(body);
      out.push(fill ? colorName(fill[1]) : 'red');
    }
    return out;
  }

  function colorName(hex) {
    const h = hex.slice(-6);
    const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
    if (g > r + 10 && g > b) return 'green';
    if (r > 200 && g > 180 && b < 170) return 'yellow';
    return 'red';
  }

  const CF_OPS = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=', equal: '=', notEqual: '<>', between: 'between' };

  function cfValue(f) {
    const t = unesc(f || '').trim();
    if (/^".*"$/.test(t)) return t.slice(1, -1).replace(/""/g, '"');
    return t !== '' && !Number.isNaN(Number(t)) ? Number(t) : t;
  }

  function readCondFormats(ctx, p, dxfStyles) {
    const s = p.sheet;
    s.condFormats = s.condFormats || [];
    const re = /<(?:\w+:)?conditionalFormatting\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?conditionalFormatting>/g;
    let m;
    while ((m = re.exec(p.xml))) {
      const a = attrs('<x ' + m[1] + '>');
      if (!a.sqref) continue; // x14 extension blocks (data bars, icon sets) are skipped
      const rules = m[2].match(/<(?:\w+:)?cfRule\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?cfRule>)/g) || [];
      for (const raw of rules) {
        const ra = attrs(/<(?:\w+:)?cfRule\b[^>]*>/.exec(raw)[0]);
        const formulas = [];
        const fre = /<(?:\w+:)?formula>([\s\S]*?)<\/(?:\w+:)?formula>/g;
        let fm;
        while ((fm = fre.exec(raw))) formulas.push(fm[1]);
        const style = ra.dxfId !== undefined && dxfStyles[Number(ra.dxfId)] ? dxfStyles[Number(ra.dxfId)] : 'red';
        let rule = null;
        if (ra.type === 'cellIs' && CF_OPS[ra.operator]) {
          rule = { op: CF_OPS[ra.operator], value: cfValue(formulas[0]) };
          if (rule.op === 'between') rule.value2 = cfValue(formulas[1]);
        } else if (ra.type === 'containsText') rule = { op: 'contains', value: ra.text || '' };
        else if (ra.type === 'duplicateValues') rule = { op: 'duplicate' };
        else if (ra.type === 'top10' && ra.bottom !== '1' && ra.percent !== '1') rule = { op: 'top', value: Number(ra.rank || 10) };
        else if (ra.type === 'aboveAverage' && ra.aboveAverage !== '0') rule = { op: 'aboveAverage' };
        else if (ra.type === 'containsBlanks') rule = { op: 'blank' };
        if (!rule) continue;
        for (const range of a.sqref.split(/\s+/).filter(Boolean)) s.condFormats.push({ range, ...rule, style });
      }
    }
  }

  const CHART_TYPES = { barChart: null, bar3DChart: null, lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'pie', scatterChart: 'scatter', areaChart: 'area', area3DChart: 'area' };

  function richText(fragment) {
    let s = '';
    const re = /<a:t>([\s\S]*?)<\/a:t>/g;
    let m;
    while ((m = re.exec(fragment))) s += unesc(m[1]);
    return s;
  }

  function readCharts(ctx, p) {
    for (const drawing of relsOfType(ctx.files, p.path, 'drawing')) {
      for (const chartPath of relsOfType(ctx.files, drawing, 'chart')) {
        const rawXml = readText(ctx.files, chartPath);
        if (!rawXml) continue;
        // Excel prefixes chart elements with c:, other writers use a default namespace
        const xml = rawXml.replace(/<(\/?)c:/g, '<$1');
        const plot = /<plotArea>([\s\S]*?)<\/plotArea>/.exec(xml);
        if (!plot) continue;
        const km = /<(barChart|bar3DChart|lineChart|line3DChart|pieChart|pie3DChart|doughnutChart|scatterChart|areaChart|area3DChart)>([\s\S]*?)<\/\1>/.exec(plot[1]);
        if (!km) continue;
        let type = CHART_TYPES[km[1]];
        if (!type) {
          const dir = /<barDir val="(\w+)"/.exec(km[2]);
          type = dir && dir[1] === 'bar' ? 'bar' : 'column';
        }
        // data range: bounding box of every series reference on the data sheet
        let dataSheet = null, box = null;
        const fre = /<f>([\s\S]*?)<\/f>/g;
        let fm;
        while ((fm = fre.exec(km[2]))) {
          for (const piece of unesc(fm[1]).replace(/^\(|\)$/g, '').split(',')) {
            const ref = splitRef(piece);
            const r = F.parseRangeAddress(ref.range);
            if (!r) continue;
            if (dataSheet === null) dataSheet = ref.sheet || p.sheet.name;
            if ((ref.sheet || p.sheet.name) !== dataSheet) continue;
            box = box ? { r1: Math.min(box.r1, r.r1), c1: Math.min(box.c1, r.c1), r2: Math.max(box.r2, r.r2), c2: Math.max(box.c2, r.c2) } : { r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 };
          }
        }
        // title: explicit rich text, a linked cell, or (for an empty title) the first series name
        let title = '';
        const deleted = /<autoTitleDeleted val="1"\/>/.test(xml);
        const tm = /<chart>[\s\S]*?<title>([\s\S]*?)<\/title>/.exec(xml.split('<plotArea>')[0] + '<plotArea>');
        if (tm) {
          title = richText(tm[1]);
          if (!title) { const v = /<v>([\s\S]*?)<\/v>/.exec(tm[1]); if (v) title = unesc(v[1]); }
          if (!title) { const sv = /<tx>[\s\S]*?<v>([\s\S]*?)<\/v>/.exec(km[2]); if (sv) title = unesc(sv[1]); }
        } else if (!deleted && (km[2].match(/<ser>/g) || []).length === 1 && type !== 'scatter') {
          const sv = /<tx>[\s\S]*?<v>([\s\S]*?)<\/v>/.exec(km[2]);
          if (sv) title = unesc(sv[1]);
        }
        const minm = /<valAx>[\s\S]*?<scaling>[\s\S]*?<min val="([^"]+)"/.exec(plot[1]);
        const chart = {
          id: 'chart' + (ctx.wb.charts.length + 1) + '_x', sheet: p.sheet.name, type,
          range: box ? F.rangeToAddress(box) : 'A1', title, axisMin: minm ? Number(minm[1]) : null,
        };
        if (dataSheet && dataSheet !== p.sheet.name) chart.dataSheet = dataSheet;
        if (/<pivotSource>/.test(xml)) chart.pivot = true;
        ctx.wb.charts.push(chart);
      }
    }
  }

  const PIVOT_AGG = { sum: 'sum', count: 'count', average: 'average', max: 'max', min: 'min', countNums: 'count', product: 'sum' };
  const PIVOT_GROUP = { months: 'month', quarters: 'quarter', years: 'year' };

  function readPivots(ctx, p) {
    for (const ptPath of relsOfType(ctx.files, p.path, 'pivotTable')) {
      const xml = readText(ctx.files, ptPath);
      const cachePath = relsOfType(ctx.files, ptPath, 'pivotCacheDefinition')[0];
      const cache = cachePath && readText(ctx.files, cachePath);
      if (!xml || !cache) continue;
      const fields = [];
      const fre = /<(?:\w+:)?cacheField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?cacheField>)/g;
      let fm;
      while ((fm = fre.exec(cache))) {
        const a = attrs('<x ' + fm[1] + '>');
        const body = fm[2] || '';
        const fg = /<(?:\w+:)?fieldGroup\b([^>]*)>/.exec(body);
        const rp = /<(?:\w+:)?rangePr\b([^>]*)>/.exec(body);
        const g = fg ? attrs('<x ' + fg[1] + '>') : {};
        fields.push({ name: a.name, base: g.base !== undefined ? Number(g.base) : null, groupBy: rp ? attrs('<x ' + rp[1] + '>').groupBy : null });
      }
      const resolve = (x) => {
        const f = fields[x];
        if (!f) return null;
        const baseName = f.base !== null && fields[f.base] ? fields[f.base].name : f.name;
        return { name: baseName, group: PIVOT_GROUP[f.groupBy] || null };
      };
      const listOf = (tag) => {
        const blk = new RegExp('<(?:\\w+:)?' + tag + '\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?' + tag + '>').exec(xml);
        return blk ? (blk[1].match(/<(?:\w+:)?field\b[^>]*>/g) || []).map((t) => Number(attrs(t).x)).filter((x) => x >= 0) : [];
      };
      const pickField = (xs) => {
        const rs = xs.map(resolve).filter(Boolean);
        if (!rs.length) return { name: null, group: null };
        const name = rs[0].name;
        const order = ['month', 'quarter', 'year'];
        const groups = rs.filter((r) => r.name === name && r.group).map((r) => r.group).sort((a, b) => order.indexOf(a) - order.indexOf(b));
        return { name, group: groups[0] || null };
      };
      const row = pickField(listOf('rowFields'));
      const col = pickField(listOf('colFields'));
      const dm = /<(?:\w+:)?dataField\b[^>]*>/.exec(xml);
      const da = dm ? attrs(dm[0]) : null;
      const src = /<(?:\w+:)?worksheetSource\b[^>]*>/.exec(cache);
      const sa = src ? attrs(src[0]) : {};
      let sourceSheet = sa.sheet || null, sourceRange = sa.ref || null;
      if (sa.name) {
        const t = ctx.wb.tables.find((x) => x.name.toLowerCase() === sa.name.toLowerCase());
        if (t) { sourceSheet = t.sheet; sourceRange = t.ref; }
        else if (ctx.wb.names[sa.name]) { const r = splitRef(ctx.wb.names[sa.name]); sourceSheet = r.sheet; sourceRange = r.range; }
      }
      const loc = /<(?:\w+:)?location\b[^>]*>/.exec(xml);
      const pivot = {
        id: 'pivot' + (ctx.wb.pivots.length + 1) + '_x', imported: true,
        sourceSheet, sourceRange, targetSheet: p.sheet.name,
        targetCell: loc ? String(attrs(loc[0]).ref || 'A1').split(':')[0] : 'A1',
        rows: row.name, columns: col.name,
        values: da && fields[Number(da.fld)] ? fields[Number(da.fld)].name : null,
        agg: PIVOT_AGG[(da && da.subtotal) || 'sum'] || 'sum',
      };
      if (row.group) pivot.rowsGroup = row.group;
      if (col.group) pivot.columnsGroup = col.group;
      ctx.wb.pivots.push(pivot);
    }
  }

  // ----- Power Query (the DataMashup package inside customXml) -----------------

  function base64ToBytes(b64) {
    const clean = b64.replace(/\s+/g, '');
    if (typeof atob === 'function') {
      const bin = atob(clean);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    return new Uint8Array(Buffer.from(clean, 'base64'));
  }

  function decodeXmlBytes(bytes) {
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
    return dec.decode(bytes).replace(/^\uFEFF/, '');
  }

  // Returns the text of Formulas/Section1.m, or null when the file has no queries.
  async function readMashup(files) {
    for (const name of Object.keys(files)) {
      if (!/^customXml\/item\d*\.xml$/i.test(name)) continue;
      const text = decodeXmlBytes(files[name]);
      const m = /<DataMashup\b[^>]*>([\s\S]*?)<\/DataMashup>/.exec(text);
      if (!m) continue;
      const bin = base64ToBytes(m[1]);
      const dv = new DataView(bin.buffer, bin.byteOffset, bin.byteLength);
      const len = dv.getUint32(4, true);
      const pkg = await unzip(bin.subarray(8, 8 + len));
      const key = Object.keys(pkg).find((k) => /^Formulas\/Section1\.m$/i.test(k));
      return key ? dec.decode(pkg[key]) : null;
    }
    return null;
  }

  // Split on a separator at bracket depth 0, outside "strings" (M doubles quotes).
  function splitTop(text, sep) {
    const out = [];
    let depth = 0, inStr = false, start = 0;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inStr) {
        if (ch === '"') { if (text[i + 1] === '"') i++; else inStr = false; }
        continue;
      }
      if (ch === '"') inStr = true;
      else if (ch === '(' || ch === '[' || ch === '{') depth++;
      else if (ch === ')' || ch === ']' || ch === '}') depth--;
      else if (depth === 0 && text.startsWith(sep, i)) { out.push(text.slice(start, i)); start = i + sep.length; i += sep.length - 1; }
    }
    out.push(text.slice(start));
    return out.map((x) => x.trim()).filter((x) => x !== '');
  }

  // Remove // and /* */ comments outside strings.
  function stripMComments(text) {
    let out = '', inStr = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inStr) { out += ch; if (ch === '"') { if (text[i + 1] === '"') { out += '"'; i++; } else inStr = false; } continue; }
      if (ch === '"') { inStr = true; out += ch; continue; }
      if (ch === '/' && text[i + 1] === '/') { while (i < text.length && text[i] !== '\n') i++; out += '\n'; continue; }
      if (ch === '/' && text[i + 1] === '*') { const e = text.indexOf('*/', i + 2); i = e < 0 ? text.length : e + 1; out += ' '; continue; }
      out += ch;
    }
    return out;
  }

  const mStr = (s) => { const t = String(s || '').trim(); return /^".*"$/s.test(t) ? t.slice(1, -1).replace(/""/g, '"') : t; };
  const mIdent = (s) => { const t = String(s || '').trim(); return /^#".*"$/s.test(t) ? mStr(t.slice(1)) : t; };
  const mList = (s) => { const t = String(s || '').trim(); return t.startsWith('{') && t.endsWith('}') ? splitTop(t.slice(1, -1), ',') : [t]; };
  const mCol = (s) => { const m = /^\[\s*(#"(?:[^"]|"")*"|[^\]]+?)\s*\]$/.exec(String(s).trim()); return m ? mIdent(m[1]) : null; };

  // "Table.Foo(a, b)" -> { fn: 'Table.Foo', args: ['a', 'b'] }
  function mCall(expr) {
    const m = /^([A-Za-z][\w.]*)\s*\(/.exec(expr.trim());
    if (!m) return null;
    const t = expr.trim();
    let depth = 0, inStr = false;
    for (let i = m[0].length - 1; i < t.length; i++) {
      const ch = t[i];
      if (inStr) { if (ch === '"') { if (t[i + 1] === '"') i++; else inStr = false; } continue; }
      if (ch === '"') inStr = true;
      else if (ch === '(' || ch === '[' || ch === '{') depth++;
      else if (ch === ')' || ch === ']' || ch === '}') { depth--; if (depth === 0) return { fn: m[1], args: splitTop(t.slice(m[0].length, i), ','), rest: t.slice(i + 1).trim() }; }
    }
    return null;
  }

  // Split "let a = x, b = y in b" into [{ name, expr }].
  function mLet(body) {
    const t = body.trim();
    if (!/^let\b/.test(t)) return [{ name: 'Source', expr: t }];
    let depth = 0, inStr = false, lets = 0, inAt = -1;
    for (let i = 0; i < t.length; i++) {
      const ch = t[i];
      if (inStr) { if (ch === '"') { if (t[i + 1] === '"') i++; else inStr = false; } continue; }
      if (ch === '"') { inStr = true; continue; }
      if ('([{'.includes(ch)) depth++;
      else if (')]}'.includes(ch)) depth--;
      else if (depth === 0 && /[A-Za-z]/.test(ch) && !/[\w.#]/.test(t[i - 1] || ' ')) {
        const w = /^[A-Za-z]+/.exec(t.slice(i))[0];
        if (w === 'let') lets++;
        if (w === 'in') { lets--; if (lets === 0) { inAt = i; break; } }
        i += w.length - 1;
      }
    }
    const inner = t.slice(3, inAt < 0 ? t.length : inAt);
    return splitTop(inner, ',').map((part) => {
      const m = /^(#"(?:[^"]|"")*"|[A-Za-z_][\w.]*)\s*=\s*([\s\S]*)$/.exec(part);
      return m ? { name: mIdent(m[1]), expr: m[2].trim() } : null;
    }).filter(Boolean);
  }

  const M_TYPES = { 'type number': 'number', 'type text': 'text', 'type date': 'date', 'type datetime': 'date', 'type datetimezone': 'date', 'type logical': 'boolean', 'Int64.Type': 'whole', 'Currency.Type': 'currency', 'Percentage.Type': 'number', 'type any': 'any' };
  const M_TEXT_FNS = { 'Text.Trim': 'trim', 'Text.Clean': 'clean', 'Text.Upper': 'upper', 'Text.Lower': 'lower', 'Text.Proper': 'proper' };
  const M_AGG = { 'List.Sum': 'sum', 'List.Average': 'average', 'List.Min': 'min', 'List.Max': 'max', 'Table.RowCount': 'count', 'List.Count': 'count' };

  // "(x)" -> "x" only when the parentheses wrap the whole expression.
  function unwrapParens(text) {
    let t = text.trim();
    while (t.startsWith('(') && t.endsWith(')')) {
      let depth = 0, inStr = false, closesAtEnd = true;
      for (let i = 0; i < t.length; i++) {
        const ch = t[i];
        if (inStr) { if (ch === '"') { if (t[i + 1] === '"') i++; else inStr = false; } continue; }
        if (ch === '"') inStr = true;
        else if (ch === '(') depth++;
        else if (ch === ')') { depth--; if (depth === 0 && i < t.length - 1) { closesAtEnd = false; break; } }
      }
      if (!closesAtEnd) break;
      t = t.slice(1, -1).trim();
    }
    return t;
  }

  function mFilter(cond) {
    let c = unwrapParens(cond.trim().replace(/^each\s+/, ''));
    if (/List\.RemoveMatchingItems\s*\(\s*Record\.FieldValues/.test(c)) return { type: 'removeBlankRows' };
    const neg = /^not\s+/.test(c);
    if (neg) c = c.replace(/^not\s+/, '');
    const fn = mCall(c);
    if (fn && /^Text\.(Contains|StartsWith|EndsWith)$/.test(fn.fn)) {
      const op = { 'Text.Contains': neg ? 'notContains' : 'contains', 'Text.StartsWith': 'beginsWith', 'Text.EndsWith': 'endsWith' }[fn.fn];
      return { type: 'filterRows', column: mCol(fn.args[0]), op, value: mStr(fn.args[1]) };
    }
    const first = splitTop(c, ' and ')[0];
    const m = /^(\[[^\]]+\])\s*(<>|>=|<=|=|>|<)\s*([\s\S]+)$/.exec(first);
    if (m) {
      const raw = m[3].trim();
      if (raw === 'null' || raw === '""') return { type: 'filterRows', column: mCol(m[1]), op: m[2] === '=' ? 'blank' : 'notBlank' };
      const v = /^".*"$/.test(raw) ? mStr(raw) : !Number.isNaN(Number(raw)) ? Number(raw) : raw;
      return { type: 'filterRows', column: mCol(m[1]), op: m[2], value: v };
    }
    return { type: 'filterRows', column: null, op: 'custom', value: cond };
  }

  // One M step -> zero or more of our steps (types matter most for grading).
  function mSteps(expr) {
    const call = mCall(expr);
    if (!call) return [];
    const a = call.args;
    switch (call.fn) {
      case 'Table.TransformColumnTypes': {
        const columns = {};
        mList(a[1]).forEach((pair) => { const [c, t] = mList(pair); if (c) columns[mStr(c)] = M_TYPES[String(t || '').trim()] || 'any'; });
        return [{ type: 'changeType', columns }];
      }
      case 'Table.RemoveColumns': return [{ type: 'removeColumns', columns: mList(a[1]).map(mStr) }];
      case 'Table.SelectColumns': return [{ type: 'keepColumns', columns: mList(a[1]).map(mStr) }];
      case 'Table.ReorderColumns': return [];
      case 'Table.RenameColumns': {
        const pairs = mList(a[1]);
        const list = pairs.length && pairs[0].trim().startsWith('{') ? pairs : [a[1]];
        return list.map((pair) => { const [f, t] = mList(pair); return { type: 'renameColumn', from: mStr(f), to: mStr(t) }; });
      }
      case 'Table.SelectRows': return [mFilter(a[1] || '')];
      case 'Table.ReplaceValue': {
        const cols = mList(a[4]).map(mStr);
        return (cols.length ? cols : [null]).map((column) => ({ type: 'replaceValues', column, find: mStr(a[1]) === 'null' ? '' : mStr(a[1]), replace: mStr(a[2]) === 'null' ? '' : mStr(a[2]), matchEntire: /ReplaceValue$/.test(String(a[3] || '')) }));
      }
      case 'Table.TransformColumns': {
        const items = mList(a[1]);
        const list = items.length && items[0].trim().startsWith('{') ? items : [a[1]];
        return list.map((it) => { const [c, f] = mList(it); const fn = String(f || '').trim().replace(/^each\s+/, '').replace(/\(_\)$/, ''); return { type: 'transform', column: mStr(c), op: M_TEXT_FNS[fn] || 'other' }; });
      }
      case 'Table.SplitColumn': {
        const sp = mCall(a[2] || '');
        let delimiter = sp ? mStr(sp.args[0]) : ',';
        if (delimiter === '#(tab)') delimiter = 'tab';
        return [{ type: 'splitColumn', column: mStr(a[1]), delimiter, names: a[3] ? mList(a[3]).map(mStr) : undefined }];
      }
      case 'Table.Distinct': return [{ type: 'removeDuplicates', columns: a[1] ? mList(a[1]).map(mStr) : [] }];
      case 'Table.RemoveRowsWithErrors': return [{ type: 'removeErrors' }];
      case 'Table.Sort': {
        const first = mList(a[1])[0] || '';
        const parts = first.trim().startsWith('{') ? mList(first) : [first];
        return [{ type: 'sortRows', column: mStr(parts[0]), order: /Descending/.test(parts[1] || '') ? 'desc' : 'asc' }];
      }
      case 'Table.AddColumn': return [{ type: 'addColumn', name: mStr(a[1]), formula: String(a[2] || '').replace(/^each\s+/, '') }];
      case 'Table.Group': {
        const aggs = mList(a[2]).map((it) => {
          const [n, f] = mList(it);
          const fc = mCall(String(f || '').replace(/^each\s+/, ''));
          return { name: mStr(n), op: fc ? M_AGG[fc.fn] || 'count' : 'count', column: fc && fc.args[0] ? mCol(fc.args[0]) : null };
        });
        return [{ type: 'groupBy', by: mList(a[1]).map(mStr), aggregations: aggs }];
      }
      case 'Table.UnpivotOtherColumns': return [{ type: 'unpivot', keep: mList(a[1]).map(mStr), attributeName: mStr(a[2] || '"Attribute"'), valueName: mStr(a[3] || '"Value"') }];
      case 'Table.Unpivot': return [{ type: 'unpivot', columns: mList(a[1]).map(mStr), keep: [] }];
      case 'Table.NestedJoin':
      case 'Table.Join': {
        const kindArg = a.find((x) => /JoinKind\./.test(x)) || 'JoinKind.LeftOuter';
        const k = /JoinKind\.(\w+)/.exec(kindArg)[1];
        return [{ type: 'merge', with: { kind: 'query', name: mIdent(a[2]) }, leftKey: mStr(mList(a[1])[0]), rightKey: mStr(mList(a[3])[0]), kind: k[0].toLowerCase() + k.slice(1) }];
      }
      case 'Table.Combine': return [{ type: 'append', with: mList(a[0]).map((x) => ({ kind: 'query', name: mIdent(x) })) }];
      case 'Table.FirstN': return [{ type: 'keepTopRows', count: Number(a[1]) }];
      case 'Table.FillDown': return mList(a[1]).map((c) => ({ type: 'fillDown', column: mStr(c) }));
      default: return [];
    }
  }

  function mSource(text, queryNames) {
    let m = /File\.Contents\s*\(\s*"((?:[^"]|"")*)"/.exec(text);
    if (m && !/Folder\.Files/.test(text)) return { kind: 'file', file: mStr('"' + m[1] + '"').split(/[\\/]/).pop() };
    m = /Folder\.Files\s*\(\s*"((?:[^"]|"")*)"/.exec(text);
    if (m) return { kind: 'folder', files: [], path: mStr('"' + m[1] + '"') };
    m = /Excel\.CurrentWorkbook\s*\(\s*\)\s*\{\s*\[\s*Name\s*=\s*"((?:[^"]|"")*)"/.exec(text);
    if (m) return { kind: 'table', table: mStr('"' + m[1] + '"') };
    return null;
  }

  async function readQueries(ctx) {
    let m;
    try { m = await readMashup(ctx.files); } catch (e) { return; }
    if (!m) return;
    const text = stripMComments(m);
    const heads = [];
    const re = /shared\s+(#"(?:[^"]|"")*"|[A-Za-z_][\w.]*)\s*=/g;
    let hm;
    while ((hm = re.exec(text))) heads.push({ name: mIdent(hm[1]), start: hm.index, bodyStart: re.lastIndex });
    const names = heads.map((h) => h.name);
    // Map loaded tables to queries via connections.xml ("Query - Name").
    const connXml = readText(ctx.files, 'xl/connections.xml') || '';
    const conns = {};
    (connXml.match(/<(?:\w+:)?connection\b[^>]*>/g) || []).forEach((tag) => {
      const a = attrs(tag);
      const nm = /^Query - (.+)$/.exec(a.name || '');
      if (a.id) conns[a.id] = nm ? nm[1] : a.name;
    });
    const wb = ctx.wb;
    heads.forEach((h, i) => {
      const body = text.slice(h.bodyStart, i + 1 < heads.length ? heads[i + 1].start : text.length).trim().replace(/;\s*$/, '');
      const lets = mLet(body);
      let source = mSource(body, names);
      if (!source && lets.length) {
        const first = mIdent(lets[0].expr);
        const ref = names.find((n) => n.toLowerCase() === first.toLowerCase());
        if (ref) source = { kind: 'query', name: ref };
      }
      if (source && source.kind === 'table') {
        const t = wb.tables.find((x) => x.name.toLowerCase() === source.table.toLowerCase());
        source = t ? { kind: 'sheet', sheet: t.sheet, range: t.ref, table: t.name } : { kind: 'sheet', sheet: source.table };
      }
      const steps = [];
      lets.forEach((l) => { steps.push(...mSteps(l.expr)); });
      const def = { id: 'q' + (wb.queries.length + 1) + '_x', name: h.name, source: source || { kind: 'query', name: '?' }, steps, imported: true, m: body };
      // where did it load?
      const norm = (s) => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
      const t = wb.tables.find((x) => x.connectionId && norm(conns[x.connectionId] || '') === norm(h.name))
        || wb.tables.find((x) => norm(x.name) === norm(h.name));
      const linked = wb.tables.some((x) => x.connectionId);
      if (t) { def.loadedSheet = t.sheet; def.loadedRange = t.ref; }
      // files without query-table links: fall back to a sheet named after the query
      else if (!linked && wb.getSheet(h.name) && !(def.source.kind === 'sheet' && def.source.sheet === h.name)) def.loadedSheet = h.name;
      if (def.loadedSheet) {
        const rec = wb.recordsFrom(def.loadedSheet, def.loadedRange);
        if (!rec.error) def.loadedRows = rec.rows.length;
      }
      wb.queries.push(def);
    });
  }

  // Browser helper: trigger a download of bytes.
  function download(bytes, filename, mime) {
    const blob = new Blob([bytes], { type: mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  const api = { write, read, unzip, zipStore, crc32, download, namedFormatFor, readMashup, mLet, mSteps, mSource, splitTop };
  root.DIM = root.DIM || {};
  root.DIM.Xlsx = api;
  if (isNode) module.exports = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
