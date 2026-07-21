import assert from 'node:assert/strict'
import { test } from 'node:test'
import { type CellInput, patchSheet } from './patch.js'

const sheet = (rows: string) =>
  `<?xml version="1.0"?><worksheet xmlns="http://x"><cols><col min="1" max="1" width="20"/></cols>` +
  `<sheetData>${rows}</sheetData>` +
  `<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><pageMargins left="0.7"/></worksheet>`

const ROWS =
  '<row r="1" ht="30" customHeight="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="4"><v>1</v></c></row>' +
  '<row r="2" hidden="1"><c r="A2"><v>2</v></c></row>'

const patch = (edits: Array<[string, CellInput]>) => patchSheet(sheet(ROWS), new Map(edits), false)

test('returns the source untouched when there is nothing to do', () => {
  const source = sheet(ROWS)

  assert.equal(patchSheet(source, new Map(), false), source)
})

test('replaces the value of a cell', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<c r="A2"><v>99<\/v><\/c>/)
})

test('keeps the style of the cell it replaces', () => {
  const patched = patch([['B1', 7]])

  assert.match(patched, /<c r="B1" s="4"><v>7<\/v><\/c>/)
})

test('leaves every other cell alone', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<c r="A1" t="s"><v>0<\/v><\/c>/)
  assert.match(patched, /<c r="B1" s="4"><v>1<\/v><\/c>/)
})

test('leaves row attributes alone', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<row r="1" ht="30" customHeight="1">/)
  assert.match(patched, /<row r="2" hidden="1">/)
})

test('leaves everything outside the cell data alone', () => {
  const patched = patch([['A2', 99]])

  assert.match(patched, /<col min="1" max="1" width="20"\/>/)
  assert.match(patched, /<mergeCell ref="A1:B1"\/>/)
  assert.match(patched, /<pageMargins left="0.7"\/>/)
})

test('replaces a cell that was written self closing', () => {
  const source = sheet('<row r="1"><c r="A1" s="2"/></row>')

  const patched = patchSheet(source, new Map<string, CellInput>([['A1', 5]]), false)

  assert.match(patched, /<c r="A1" s="2"><v>5<\/v><\/c>/)
})

test('writes text as an inline string', () => {
  const patched = patch([['A2', 'hello']])

  assert.match(patched, /<c r="A2" t="inlineStr"><is><t>hello<\/t><\/is><\/c>/)
})

test('escapes text that would break the markup', () => {
  const patched = patch([['A2', 'a & b < c']])

  assert.match(patched, /<t>a &amp; b &lt; c<\/t>/)
})

test('writes a boolean', () => {
  const patched = patch([['A2', true]])

  assert.match(patched, /<c r="A2" t="b"><v>1<\/v><\/c>/)
})

test('writes null as a cell with no value', () => {
  const patched = patch([['B1', null]])

  assert.match(patched, /<c r="B1" s="4"\/>/)
})

test('writes a date as its serial', () => {
  const patched = patch([['A2', new Date('2024-01-01T00:00:00Z')]])

  assert.match(patched, /<c r="A2"><v>45292<\/v><\/c>/)
})

test('writes a date against the workbook epoch', () => {
  const patched = patchSheet(
    sheet(ROWS),
    new Map<string, CellInput>([['A2', new Date('1904-01-01T00:00:00Z')]]),
    true,
  )

  assert.match(patched, /<c r="A2"><v>0<\/v><\/c>/)
})

test('applies several edits at once', () => {
  const patched = patch([
    ['A2', 1],
    ['B1', 2],
  ])

  assert.match(patched, /<c r="B1" s="4"><v>2<\/v><\/c>/)
  assert.match(patched, /<c r="A2"><v>1<\/v><\/c>/)
})

test('adds a cell that was not in the sheet yet', () => {
  const patched = patch([['Z9', 1]])

  assert.match(patched, /<row r="9"><c r="Z9"><v>1<\/v><\/c><\/row>/)
})

test('rejects a number that cannot be written', () => {
  assert.throws(() => patch([['A2', Number.POSITIVE_INFINITY]]), /cannot hold/)
})

test('keeps whitespace on an inline string', () => {
  const patched = patch([['A2', '  padded  ']])

  assert.match(patched, /<t xml:space="preserve"> {2}padded {2}<\/t>/)
})

test('rejects text that xml cannot represent', () => {
  assert.throws(() => patch([['A2', `a${String.fromCharCode(7)}b`]]), /cannot be written to xml/i)
})
