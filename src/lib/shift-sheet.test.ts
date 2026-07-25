import assert from 'node:assert/strict'
import { test } from 'node:test'

import type { ShiftSpec } from './shift.js'
import { shiftForeignFormulas, shiftSheet } from './shift-sheet.js'

const insertAt = (at: number): ShiftSpec => ({
  axis: 'row',
  at,
  delta: 1,
  editedSheet: 'Sheet1',
  onCurrentSheet: true,
})

const insertColumnAt = (at: number): ShiftSpec => ({
  axis: 'column',
  at,
  delta: 1,
  editedSheet: 'Sheet1',
  onCurrentSheet: true,
})

const SHEET =
  '<worksheet><dimension ref="A1:B4"/><sheetData>' +
  '<row r="1"><c r="A1"><v>1</v></c></row>' +
  '<row r="2"><c r="A2"><f>A1+1</f><v>2</v></c></row>' +
  '<row r="4"><c r="B4"><f>SUM(A1:A2)</f></c></row>' +
  '</sheetData><mergeCells count="1"><mergeCell ref="A4:B4"/></mergeCells></worksheet>'

test('shiftSheet renumbers rows and cells at or past the insertion point', () => {
  const out = shiftSheet(SHEET, insertAt(2))
  assert.match(out, /<row r="1"><c r="A1">/)
  assert.match(out, /<row r="3"><c r="A3">/)
  assert.match(out, /<row r="5"><c r="B5">/)
})

test('shiftSheet rewrites formula text, merges and the dimension', () => {
  const out = shiftSheet(SHEET, insertAt(2))
  assert.match(out, /<f>A1\+1<\/f>/)
  assert.match(out, /<f>SUM\(A1:A3\)<\/f>/)
  assert.match(out, /<mergeCell ref="A5:B5"\/>/)
  assert.match(out, /<dimension ref="A1:B5"\/>/)
})

test('shiftSheet re-encodes an entity in a shifted formula', () => {
  const sheet = '<sheetData><row r="4"><c r="A4"><f>A1&lt;A5</f></c></row></sheetData>'
  assert.match(shiftSheet(sheet, insertAt(2)), /<f>A1&lt;A6<\/f>/)
})

test('shiftSheet shifts sqref and ref on conditional formatting and hyperlinks', () => {
  const sheet =
    '<worksheet><sheetData><row r="5"><c r="A5"/></row></sheetData>' +
    '<conditionalFormatting sqref="A5:A9"><cfRule/></conditionalFormatting>' +
    '<hyperlinks><hyperlink ref="A5" r:id="rId1"/></hyperlinks></worksheet>'
  const out = shiftSheet(sheet, insertAt(3))
  assert.match(out, /<conditionalFormatting sqref="A6:A10">/)
  assert.match(out, /<hyperlink ref="A6" r:id="rId1"\/>/)
})

test('shiftSheet shifts a shared formula master range and its text, leaving slaves', () => {
  const sheet =
    '<sheetData><row r="5"><c r="B5"><f t="shared" ref="B5:B10" si="0">A5*2</f></c></row>' +
    '<row r="6"><c r="B6"><f t="shared" si="0"/></c></row></sheetData>'
  const out = shiftSheet(sheet, insertAt(3))
  assert.match(out, /<f t="shared" ref="B6:B11" si="0">A6\*2<\/f>/)
  assert.match(out, /<row r="7"><c r="B7"><f t="shared" si="0"\/><\/c><\/row>/)
})

test('shiftSheet numbers an implicit row and makes the shift explicit', () => {
  const sheet = '<sheetData><row r="4"><c r="A4"/></row><row><c r="A5"/></row></sheetData>'
  const out = shiftSheet(sheet, insertAt(3))
  assert.match(out, /<row r="5"><c r="A5"\/><\/row><row r="6"><c r="A6"\/><\/row>/)
})

test('shiftSheet shifts columns, cols bounds and cell columns without renumbering rows', () => {
  const sheet =
    '<worksheet><cols><col min="2" max="2" width="9"/><col min="1" max="5" width="4"/></cols>' +
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><f>A1+B1</f></c></row></sheetData>' +
    '<mergeCells count="1"><mergeCell ref="B1:C1"/></mergeCells></worksheet>'
  const out = shiftSheet(sheet, insertColumnAt(2))
  assert.match(out, /<col min="3" max="3" width="9"\/>/)
  assert.match(out, /<col min="1" max="6" width="4"\/>/)
  assert.match(out, /<row r="1"><c r="A1"><v>1<\/v><\/c><c r="D1"><f>A1\+C1<\/f>/)
  assert.match(out, /<mergeCell ref="C1:D1"\/>/)
})

test('shiftSheet leaves a sheet with nothing at or past the point unchanged', () => {
  const sheet = '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
  assert.equal(shiftSheet(sheet, insertAt(5)), sheet)
})

test('shiftForeignFormulas shifts only qualified references, leaving rows and cells put', () => {
  const foreign = { ...insertAt(2), onCurrentSheet: false }
  const sheet =
    '<sheetData><row r="8"><c r="A8"><f>Sheet1!A5+A8</f><v>7</v></c>' +
    '<c r="B8"><f t="shared" si="0"/></c></row></sheetData>'
  const out = shiftForeignFormulas(sheet, foreign)
  assert.match(out, /<row r="8"><c r="A8">/)
  assert.match(out, /<f>Sheet1!A6\+A8<\/f>/)
  assert.match(out, /<c r="B8"><f t="shared" si="0"\/><\/c>/)
})
