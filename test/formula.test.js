'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { F, wb } = require('./helpers');

// A small sheet most tests read from.
const DATA = [
  ['Region', 'Product', 'Units', 'Price', 'Date'],
  ['East', 'Pens', 10, 2.5, '2026-01-05'],
  ['West', 'Paper', 4, 3, '2026-02-10'],
  ['East', 'Paper', 7, 3, '2026-02-14'],
  ['North', 'Pens', 12, 2.5, '2026-03-01'],
  ['West', 'Ink', 3, 9.99, '2026-03-20'],
];

function calc(formula, extra = {}) {
  const book = wb({ Data: DATA, ...extra });
  book.setInput('Data', 'H1', formula, { force: true });
  return book.getValue('Data', 'H1');
}
const code = (v) => (F.isError(v) ? v.code : v);

test('addresses round-trip between A1 text and row/col indexes', () => {
  assert.deepEqual(F.parseAddress('B3'), { row: 2, col: 1, colAbs: false, rowAbs: false });
  assert.deepEqual(F.parseAddress('$C$7'), { row: 6, col: 2, colAbs: true, rowAbs: true });
  assert.equal(F.colToIndex('A'), 0);
  assert.equal(F.colToIndex('AA'), 26);
  assert.equal(F.indexToCol(27), 'AB');
  assert.equal(F.toAddress(9, 2), 'C10');
  assert.equal(F.rangeToAddress(F.parseRangeAddress('b2:d5')), 'B2:D5');
});

test('arithmetic follows Excel precedence', () => {
  assert.equal(calc('=1+2*3'), 7);
  assert.equal(calc('=(1+2)*3'), 9);
  assert.equal(calc('=2^3^2'), 64); // Excel evaluates ^ left to right
  assert.equal(calc('=-2^2'), 4); // unary minus binds tighter than ^ in Excel
  assert.equal(calc('=10%'), 0.1);
  assert.equal(calc('="a"&"b"&1'), 'ab1');
  assert.equal(code(calc('=1/0')), '#DIV/0!');
});

test('aggregates skip text and blanks like Excel', () => {
  assert.equal(calc('=SUM(C2:C6)'), 36);
  assert.equal(calc('=AVERAGE(C2:C6)'), 7.2);
  assert.equal(calc('=COUNT(A1:E6)'), 15); // dates are numbers too
  assert.equal(calc('=COUNTA(A1:A6)'), 6);
  assert.equal(calc('=COUNTBLANK(A1:A8)'), 2);
  assert.equal(calc('=MAX(C:C)'), 12);
  assert.equal(calc('=MIN(C2:C6)'), 3);
  assert.equal(calc('=MEDIAN(C2:C6)'), 7);
  assert.equal(calc('=ROUND(SUMPRODUCT(C2:C6,D2:D6),2)'), 117.97);
});

test('conditional aggregates: COUNTIF(S), SUMIF(S), AVERAGEIF(S), MAXIFS', () => {
  assert.equal(calc('=COUNTIF(A2:A6,"East")'), 2);
  assert.equal(calc('=COUNTIF(C2:C6,">=7")'), 3);
  assert.equal(calc('=COUNTIF(B2:B6,"P*")'), 4);
  assert.equal(calc('=COUNTIFS(A2:A6,"East",B2:B6,"Paper")'), 1);
  assert.equal(calc('=SUMIF(A2:A6,"West",C2:C6)'), 7);
  assert.equal(calc('=SUMIFS(C2:C6,A2:A6,"<>East",D2:D6,"<5")'), 16);
  assert.equal(calc('=AVERAGEIFS(C2:C6,B2:B6,"Pens")'), 11);
  assert.equal(calc('=MAXIFS(C2:C6,A2:A6,"West")'), 4);
  assert.equal(calc('=MINIFS(C2:C6,A2:A6,"East")'), 7);
});

test('logic: IF, IFS, AND, OR, NOT, SWITCH, IFERROR', () => {
  assert.equal(calc('=IF(C2>5,"High","Low")'), 'High');
  assert.equal(calc('=IFS(C3>10,"A",C3>3,"B",TRUE,"C")'), 'B');
  assert.equal(code(calc('=IFS(C3>100,"A")')), '#N/A');
  assert.equal(calc('=AND(C2>5,D2<3)'), true);
  assert.equal(calc('=OR(C2>50,D2<1)'), false);
  assert.equal(calc('=NOT(ISBLANK(A2))'), true);
  assert.equal(calc('=SWITCH(A3,"East",1,"West",2,0)'), 2);
  assert.equal(calc('=IFERROR(1/0,"none")'), 'none');
});

test('lookups: XLOOKUP, VLOOKUP, INDEX/MATCH and their misses', () => {
  const extra = { Products: [['ID', 'Name', 'Cost'], ['P1', 'Pens', 1.1], ['P2', 'Paper', 2.2], ['P3', 'Ink', 5]] };
  assert.equal(calc('=XLOOKUP("P2",Products!A2:A4,Products!B2:B4)', extra), 'Paper');
  assert.equal(calc('=XLOOKUP("P9",Products!A2:A4,Products!B2:B4,"Not found")', extra), 'Not found');
  assert.equal(code(calc('=XLOOKUP("P9",Products!A2:A4,Products!B2:B4)', extra)), '#N/A');
  assert.equal(calc('=VLOOKUP("P3",Products!A2:C4,3,FALSE)', extra), 5);
  assert.equal(code(calc('=VLOOKUP("P3",Products!A2:C4,4,FALSE)', extra)), '#REF!');
  assert.equal(calc('=INDEX(Products!B2:B4,MATCH("P1",Products!A2:A4,0))', extra), 'Pens');
  assert.equal(calc('=XMATCH("Ink",Products!B2:B4)', extra), 3);
});

test('text functions clean and reshape strings', () => {
  assert.equal(calc('=TRIM("  too   many  spaces ")'), 'too many spaces');
  assert.equal(calc('=PROPER("mARIA lopez")'), 'Maria Lopez');
  assert.equal(calc('=UPPER(LEFT("east",1))&MID("east",2,3)'), 'East');
  assert.equal(calc('=RIGHT("S1001",4)'), '1001');
  assert.equal(calc('=LEN("data")'), 4);
  assert.equal(calc('=FIND("@","kim@school.org")'), 4);
  assert.equal(calc('=SUBSTITUTE("555-0101","-","")'), '5550101');
  assert.equal(calc('=TEXTJOIN(", ",TRUE,A2:A4)'), 'East, West, East');
  assert.equal(calc('=CONCAT(A2,"-",B2)'), 'East-Pens');
  assert.equal(calc('=VALUE("42")'), 42);
  assert.equal(calc('=TEXT(0.256,"0.0%")'), '25.6%');
});

test('dates are serial numbers, as in Excel', () => {
  assert.equal(F.dateToSerial(2026, 1, 5), 46027);
  assert.deepEqual(F.serialToDate(46027), { y: 2026, m: 1, d: 5, dow: 1 });
  assert.equal(calc('=DATE(2026,1,5)'), 46027);
  assert.equal(calc('=YEAR(E2)'), 2026);
  assert.equal(calc('=MONTH(E3)'), 2);
  assert.equal(calc('=DAY(E5)'), 1);
  assert.equal(calc('=E3-E2'), 36);
  assert.equal(calc('=TEXT(E2,"mmm")'), 'Jan');
  assert.equal(calc('=EOMONTH(E2,0)-E2'), 26);
  assert.equal(calc('=WEEKDAY(DATE(2026,1,5))'), 2); // a Monday
});

test('error values propagate and can be detected', () => {
  assert.equal(code(calc('=SUM(1,#N/A)')), '#N/A');
  assert.equal(calc('=ISNA(NA())'), true);
  assert.equal(calc('=ISERROR(1/0)'), true);
  assert.equal(calc('=IFNA(NA(),0)'), 0);
  assert.equal(code(calc('=NOSUCHFN(1)')), '#NAME?');
  assert.equal(code(calc('="a"+1')), '#VALUE!');
});

test('circular references are reported, not looped forever', () => {
  const book = wb({ S: [[1]] });
  book.setInput('S', 'B1', '=C1+1', { force: true });
  book.setInput('S', 'C1', '=B1+1', { force: true });
  const v = book.getValue('S', 'B1');
  assert.ok(F.isError(v));
});

test('shiftFormula moves relative refs and keeps $absolute ones (fill handle)', () => {
  assert.equal(F.shiftFormula('=A1*$B$1+B$2+$C3', 2, 1), '=B3*$B$1+C$2+$C5');
  assert.equal(F.shiftFormula('=SUM(A1:A3)', 1, 0), '=SUM(A2:A4)');
});

test('number parsing accepts what students type', () => {
  assert.equal(F.parseNumberLike('1,234.50').value, 1234.5);
  assert.equal(F.parseNumberLike('$12').value, 12);
  assert.equal(F.parseNumberLike('15%').value, 0.15);
  assert.equal(F.parseNumberLike('abc'), null);
});

test('formatValue renders the named formats', () => {
  assert.equal(F.formatValue(1234.5, 'currency'), '$1,234.50');
  assert.equal(F.formatValue(0.25, 'percent'), '25%');
  assert.equal(F.formatValue(46027, 'date'), '1/5/2026');
});
