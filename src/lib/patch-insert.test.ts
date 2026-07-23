import assert from 'node:assert/strict'
import { test } from 'node:test'
import { type CellInput, patchSheet as patchSheetBytes } from './patch.js'

const encode = (text: string) => new TextEncoder().encode(text)
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes)
const patchSheet = (
  source: string,
  edits: ReadonlyMap<string, CellInput>,
  date1904: boolean,
  sharedStrings?: ReadonlyMap<string, number>,
  styleOverrides?: ReadonlyMap<string, number>,
) => decode(patchSheetBytes(encode(source), edits, date1904, sharedStrings, styleOverrides))

const sheet = (rows: string) =>
  `<?xml version="1.0"?><worksheet xmlns="http://x"><cols><col min="1" max="1" width="20"/></cols>` +
  `<sheetData>${rows}</sheetData>` +
  `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`

const apply = (rows: string, edits: Array<[string, CellInput]>) =>
  patchSheet(sheet(rows), new Map(edits), false)

test('adds a cell after the cells already in its row', () => {
  const patched = apply('<row r="1"><c r="A1"><v>1</v></c></row>', [['B1', 2]])

  assert.match(patched, /<c r="A1"><v>1<\/v><\/c><c r="B1"><v>2<\/v><\/c><\/row>/)
})

test('adds a cell before the cells already in its row', () => {
  const patched = apply('<row r="1"><c r="C1"><v>3</v></c></row>', [['A1', 1]])

  assert.match(patched, /<row r="1"><c r="A1"><v>1<\/v><\/c><c r="C1">/)
})

test('adds a cell between the cells already in its row', () => {
  const patched = apply('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>', [
    ['B1', 2],
  ])

  assert.match(patched, /<c r="A1"><v>1<\/v><\/c><c r="B1"><v>2<\/v><\/c><c r="C1">/)
})

test('keeps the attributes of the row it adds a cell to', () => {
  const patched = apply('<row r="1" ht="30" hidden="1"><c r="A1"><v>1</v></c></row>', [['B1', 2]])

  assert.match(patched, /<row r="1" ht="30" hidden="1">/)
})

test('adds a row after the rows already in the sheet', () => {
  const patched = apply('<row r="1"><c r="A1"><v>1</v></c></row>', [['A2', 2]])

  assert.match(patched, /<row r="2"><c r="A2"><v>2<\/v><\/c><\/row><\/sheetData>/)
})

test('adds a row before the rows already in the sheet', () => {
  const patched = apply('<row r="5"><c r="A5"><v>5</v></c></row>', [['A1', 1]])

  assert.match(patched, /<sheetData><row r="1"><c r="A1"><v>1<\/v><\/c><\/row><row r="5">/)
})

test('adds a row between the rows already in the sheet', () => {
  const patched = apply(
    '<row r="1"><c r="A1"><v>1</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>',
    [['A2', 2]],
  )

  assert.match(patched, /<\/row><row r="2"><c r="A2"><v>2<\/v><\/c><\/row><row r="3">/)
})

test('adds a row to a sheet with no rows', () => {
  const patched = apply('', [['B2', 7]])

  assert.match(patched, /<sheetData><row r="2"><c r="B2"><v>7<\/v><\/c><\/row><\/sheetData>/)
})

test('adds a row to a sheet whose cell data is self closing', () => {
  const source = `<worksheet><sheetData/><pageMargins left="1"/></worksheet>`

  const patched = patchSheet(source, new Map<string, CellInput>([['A1', 1]]), false)

  assert.match(patched, /<sheetData><row r="1"><c r="A1"><v>1<\/v><\/c><\/row><\/sheetData>/)
  assert.match(patched, /<pageMargins left="1"\/>/)
})

test('adds several cells and rows in one pass', () => {
  const patched = apply('<row r="2"><c r="B2"><v>2</v></c></row>', [
    ['A2', 1],
    ['C2', 3],
    ['A1', 0],
    ['A3', 9],
  ])

  assert.match(
    patched,
    /<sheetData><row r="1"><c r="A1"><v>0<\/v><\/c><\/row><row r="2"><c r="A2"><v>1<\/v><\/c><c r="B2"><v>2<\/v><\/c><c r="C2"><v>3<\/v><\/c><\/row><row r="3"><c r="A3"><v>9<\/v><\/c><\/row><\/sheetData>/,
  )
})

test('mixes replacement and insertion', () => {
  const patched = apply('<row r="1"><c r="A1"><v>1</v></c></row>', [
    ['A1', 10],
    ['B1', 20],
  ])

  assert.match(patched, /<c r="A1"><v>10<\/v><\/c><c r="B1"><v>20<\/v><\/c>/)
})

test('leaves the rest of the sheet alone when adding', () => {
  const patched = apply('<row r="1"><c r="A1"><v>1</v></c></row>', [['B1', 2]])

  assert.match(patched, /<col min="1" max="1" width="20"\/>/)
  assert.match(patched, /<mergeCell ref="A1:B1"\/>/)
})

test('adds a text cell', () => {
  const patched = apply('<row r="1"><c r="A1"><v>1</v></c></row>', [['B1', 'new']])

  assert.match(patched, /<c r="B1" t="inlineStr"><is><t>new<\/t><\/is><\/c>/)
})

test('rejects a sheet with no cell data to add to', () => {
  assert.throws(
    () => patchSheet('<worksheet/>', new Map<string, CellInput>([['A1', 1]]), false),
    /no sheetData/i,
  )
})

test('orders the cells of a row it creates', () => {
  const patched = apply('', [
    ['C1', 3],
    ['A1', 1],
    ['B1', 2],
  ])

  assert.match(
    patched,
    /<row r="1"><c r="A1"><v>1<\/v><\/c><c r="B1"><v>2<\/v><\/c><c r="C1"><v>3<\/v><\/c><\/row>/,
  )
})

test('orders the rows it creates in an empty sheet', () => {
  const source = '<worksheet><sheetData/></worksheet>'

  const patched = patchSheet(
    source,
    new Map<string, CellInput>([
      ['A3', 3],
      ['A1', 1],
    ]),
    false,
  )

  assert.match(patched, /<row r="1">.*<\/row><row r="3">.*<\/row>/)
})

test('adds a cell to a row that carries no reference', () => {
  const patched = apply('<row><c r="A1"><v>1</v></c></row>', [['B1', 2]])

  assert.match(patched, /<c r="A1"><v>1<\/v><\/c><c r="B1"><v>2<\/v><\/c>/)
})

test('adds a cell next to one that carries no reference', () => {
  const patched = apply('<row r="1"><c><v>1</v></c></row>', [['B1', 2]])

  assert.match(patched, /<c><v>1<\/v><\/c><c r="B1"><v>2<\/v><\/c>/)
})

test('adds a cell to a row written self closing', () => {
  const patched = apply('<row r="1"/>', [['A1', 1]])

  assert.match(patched, /<c r="A1"><v>1<\/v><\/c>/)
})

test('adds a cell beside one written self closing', () => {
  const patched = apply('<row r="1"><c r="A1" s="2"/></row>', [['B1', 2]])

  assert.match(patched, /<c r="A1" s="2"\/><c r="B1"><v>2<\/v><\/c>/)
})

const withDimension = (ref: string, rows: string) =>
  `<worksheet><dimension ref="${ref}"/><sheetData>${rows}</sheetData></worksheet>`

const ROW_ONE = '<row r="1"><c r="A1"><v>1</v></c></row>'

test('widens the declared dimension rightwards for a new column', () => {
  const patched = patchSheet(
    withDimension('A1:B2', ROW_ONE),
    new Map<string, CellInput>([['D1', 4]]),
    false,
  )

  assert.match(patched, /<dimension ref="A1:D2"\/>/)
})

test('widens the declared dimension downwards for a new row', () => {
  const patched = patchSheet(
    withDimension('A1:B2', ROW_ONE),
    new Map<string, CellInput>([['B9', 9]]),
    false,
  )

  assert.match(patched, /<dimension ref="A1:B9"\/>/)
})

test('leaves a dimension naming a column past the last alone instead of crashing the save', () => {
  const patched = patchSheet(
    withDimension('A1:XFE1', ROW_ONE),
    new Map<string, CellInput>([['B9', 9]]),
    false,
  )

  assert.match(patched, /<dimension ref="A1:XFE1"\/>/)
})

test('leaves the dimension alone when the edit fits inside it', () => {
  const patched = patchSheet(
    withDimension('A1:Z100', ROW_ONE),
    new Map<string, CellInput>([['A1', 2]]),
    false,
  )

  assert.match(patched, /<dimension ref="A1:Z100"\/>/)
})

test('widens a dimension written as a single cell', () => {
  const patched = patchSheet(
    withDimension('A1', ROW_ONE),
    new Map<string, CellInput>([['C3', 3]]),
    false,
  )

  assert.match(patched, /<dimension ref="A1:C3"\/>/)
})

test('grows the dimension upwards and leftwards too', () => {
  const patched = patchSheet(
    withDimension('C3:D4', '<row r="3"><c r="C3"><v>1</v></c></row>'),
    new Map<string, CellInput>([['A1', 1]]),
    false,
  )

  assert.match(patched, /<dimension ref="A1:D4"\/>/)
})

test('ignores a dimension with an empty reference', () => {
  const patched = patchSheet(
    withDimension('', ROW_ONE),
    new Map<string, CellInput>([['D1', 4]]),
    false,
  )

  assert.match(patched, /<dimension ref=""\/>/)
})

test('writes a sheet with no dimension element unchanged in that respect', () => {
  const source = `<worksheet><sheetData>${ROW_ONE}</sheetData></worksheet>`

  const patched = patchSheet(source, new Map<string, CellInput>([['D1', 4]]), false)

  assert.equal(patched.includes('<dimension'), false)
  assert.match(patched, /<c r="D1"><v>4<\/v><\/c>/)
})
