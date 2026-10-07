// Builders for the OOXML parts Excel writes for PivotTables and Power Query.
// openpyxl cannot author these, so the tests inject hand-built parts that
// follow the shapes Excel 365 saves (same element names, attributes and
// relationship types) into a workbook written by the site.
'use strict';
const { X, patchXlsx, enc } = require('./helpers');

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

function sheetPath(sheets, name) {
  const i = sheets.indexOf(name);
  if (i < 0) throw new Error('no sheet ' + name);
  return 'xl/worksheets/sheet' + (i + 1) + '.xml';
}
const relsPathOf = (part) => part.replace(/([^/]+)$/, '_rels/$1.rels');

// Append a relationship to a .rels part (creating it), returning the new Id.
function addRel(files, part, type, target) {
  const path = relsPathOf(part);
  let xml = files[path] || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PKG + '"></Relationships>';
  const ids = [...xml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  const id = 'rId' + (Math.max(0, ...ids) + 1);
  xml = xml.replace('</Relationships>', '<Relationship Id="' + id + '" Type="' + type + '" Target="' + target + '"/></Relationships>');
  files[path] = xml;
  return id;
}

async function readParts(bytes) {
  const raw = await X.unzip(bytes);
  const files = {};
  for (const [k, v] of Object.entries(raw)) files[k] = new TextDecoder().decode(v);
  return files;
}

function zipText(files) {
  return X.zipStore(Object.entries(files).map(([name, text]) => ({ name, data: typeof text === 'string' ? enc.encode(text) : text })));
}

/**
 * Add a PivotTable the way Excel stores it.
 * spec: { sheet, location: 'A3:B8', source: { sheet, ref } | { name },
 *         fields: ['Date', 'Region', 'Revenue'], rows: [indexes], cols: [indexes],
 *         data: { fld, subtotal? }, groups: [{ name, base, groupBy }] }
 */
async function addPivot(bytes, sheets, spec) {
  const files = await readParts(bytes);
  const n = Object.keys(files).filter((k) => /pivotTables\/pivotTable\d+\.xml$/.test(k)).length + 1;
  const cachePart = 'xl/pivotCache/pivotCacheDefinition' + n + '.xml';
  const tablePart = 'xl/pivotTables/pivotTable' + n + '.xml';
  const fields = spec.fields.map((name, i) => {
    const own = (spec.groups || []).find((g) => g.name === name && g.base === i);
    return '<cacheField name="' + name + '" numFmtId="0">' + (own ? '<sharedItems/><fieldGroup base="' + i + '"><rangePr groupBy="' + own.groupBy + '"/></fieldGroup>' : '<sharedItems/>') + '</cacheField>';
  });
  (spec.groups || []).filter((g) => !spec.fields.includes(g.name)).forEach((g) => {
    fields.push('<cacheField name="' + g.name + '" numFmtId="0" databaseField="0"><fieldGroup base="' + g.base + '"><rangePr groupBy="' + g.groupBy + '"/></fieldGroup></cacheField>');
  });
  const src = spec.source.name ? '<worksheetSource name="' + spec.source.name + '"/>' : '<worksheetSource ref="' + spec.source.ref + '" sheet="' + spec.source.sheet + '"/>';
  files[cachePart] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="' + NS + '" xmlns:r="' + REL + '" refreshOnLoad="1"><cacheSource type="worksheet">' + src + '</cacheSource><cacheFields count="' + fields.length + '">' + fields.join('') + '</cacheFields></pivotCacheDefinition>';
  const fieldList = (xs) => xs.map((x) => '<field x="' + x + '"/>').join('');
  files[tablePart] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="' + NS + '" name="PivotTable' + n + '" cacheId="' + n + '" dataCaption="Values"><location ref="' + spec.location + '" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
    '<pivotFields count="' + fields.length + '"/>' +
    (spec.rows && spec.rows.length ? '<rowFields count="' + spec.rows.length + '">' + fieldList(spec.rows) + '</rowFields><rowItems count="1"><i/></rowItems>' : '') +
    (spec.cols && spec.cols.length ? '<colFields count="' + spec.cols.length + '">' + fieldList(spec.cols) + '</colFields>' : '') +
    '<colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Values" fld="' + spec.data.fld + '"' + (spec.data.subtotal ? ' subtotal="' + spec.data.subtotal + '"' : '') + ' baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>';
  addRel(files, tablePart, REL + '/pivotCacheDefinition', '../pivotCache/pivotCacheDefinition' + n + '.xml');
  addRel(files, sheetPath(sheets, spec.sheet), REL + '/pivotTable', '../pivotTables/pivotTable' + n + '.xml');
  files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>',
    '<Override PartName="/' + cachePart + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>' +
    '<Override PartName="/' + tablePart + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>');
  return zipText(files);
}

// The DataMashup binary: version, package zip, then empty permission/metadata blocks.
function dataMashup(mCode) {
  const pkg = X.zipStore([
    { name: '[Content_Types].xml', data: enc.encode('<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="text/xml"/><Default Extension="m" ContentType="application/x-ms-m"/></Types>') },
    { name: 'Config/Package.xml', data: enc.encode('<?xml version="1.0" encoding="utf-8"?><Package xmlns:xsd="http://www.w3.org/2001/XMLSchema"><Version>1.0.0.0</Version><MinVersion>1.0.0.0</MinVersion><Culture>en-US</Culture></Package>') },
    { name: 'Formulas/Section1.m', data: enc.encode(mCode) },
  ]);
  const out = new Uint8Array(8 + pkg.length + 12);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, 0, true);
  dv.setUint32(4, pkg.length, true);
  out.set(pkg, 8);
  return out; // trailing permission, metadata and binding lengths are zero
}

function utf16le(text) {
  const out = new Uint8Array(2 + text.length * 2);
  out[0] = 0xff; out[1] = 0xfe;
  for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); out[2 + i * 2] = c & 0xff; out[3 + i * 2] = c >> 8; }
  return out;
}

/**
 * Add Power Query queries. loads: [{ query, sheet, ref }] become query tables
 * linked to "Query - <name>" connections, as Close & Load creates them.
 */
async function addQueries(bytes, sheets, mCode, loads = []) {
  const files = await readParts(bytes);
  const b64 = Buffer.from(dataMashup(mCode)).toString('base64');
  files['customXml/item1.xml'] = utf16le('<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">' + b64 + '</DataMashup>');
  const conns = loads.map((l, i) => '<connection id="' + (i + 1) + '" keepAlive="1" name="Query - ' + l.query + '" description="Connection to the \'' + l.query + '\' query in the workbook." type="5" refreshedVersion="8" background="1" saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=' + l.query + ';Extended Properties=&quot;&quot;" command="SELECT * FROM [' + l.query + ']"/></connection>');
  files['xl/connections.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections xmlns="' + NS + '">' + conns.join('') + '</connections>';
  addRel(files, 'xl/workbook.xml', REL + '/connections', 'connections.xml');
  let ct = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/></Types>');
  const existing = Object.keys(files).filter((k) => /^xl\/tables\/table\d+\.xml$/.test(k)).length;
  loads.forEach((l, i) => {
    const n = existing + i + 1;
    const tablePart = 'xl/tables/table' + n + '.xml';
    const qtPart = 'xl/queryTables/queryTable' + n + '.xml';
    files[tablePart] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="' + NS + '" id="' + n + '" name="' + l.query.replace(/\W/g, '_') + '" displayName="' + l.query.replace(/\W/g, '_') + '" ref="' + l.ref + '" tableType="queryTable" totalsRowShown="0"><autoFilter ref="' + l.ref + '"/><tableColumns count="1"><tableColumn id="1" uniqueName="1" name="x" queryTableFieldId="1"/></tableColumns><tableStyleInfo name="TableStyleMedium7" showRowStripes="1"/></table>';
    files[qtPart] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<queryTable xmlns="' + NS + '" name="ExternalData_' + (i + 1) + '" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0" connectionId="' + (i + 1) + '"/>';
    addRel(files, tablePart, REL + '/queryTable', '../queryTables/queryTable' + n + '.xml');
    const sp = sheetPath(sheets, l.sheet);
    const rid = addRel(files, sp, REL + '/table', '../tables/table' + n + '.xml');
    files[sp] = /<tableParts/.test(files[sp])
      ? files[sp].replace(/<tableParts count="(\d+)">/, (m, c) => '<tableParts count="' + (Number(c) + 1) + '">').replace('</tableParts>', '<tablePart r:id="' + rid + '"/></tableParts>')
      : files[sp].replace('</worksheet>', '<tableParts count="1"><tablePart r:id="' + rid + '"/></tableParts></worksheet>');
    ct = ct.replace('</Types>', '<Override PartName="/' + tablePart + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/' + qtPart + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>');
  });
  files['[Content_Types].xml'] = ct;
  return zipText(files);
}

module.exports = { addPivot, addQueries, dataMashup, patchXlsx };
