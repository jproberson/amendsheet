import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  type AddedSheet,
  checkSheetName,
  withSheetContentTypes,
  withSheetRelationships,
  withSheetRemoved,
  withSheetRenamed,
  withSheetsAdded,
} from './add-sheet.js'
import { XlsxError } from './errors.js'

const added = (name: string, path: string): AddedSheet => ({
  reference: { name, path, sheetId: '5', state: 'visible' },
  relationshipId: 'rId9',
  target: path.replace(/^xl\//, ''),
})
const one = added('New', 'xl/worksheets/sheet3.xml')

const isUnwritable = (error: unknown) =>
  error instanceof XlsxError && error.code === 'unwritable-value'
const isInvalid = (error: unknown) => error instanceof XlsxError && error.code === 'invalid-content'

test('withSheetsAdded inserts before the closing sheets tag', () => {
  const out = withSheetsAdded(
    '<workbook><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>',
    [one],
  )
  assert.match(out, /<sheet name="New" sheetId="5" r:id="rId9"\/><\/sheets>/)
})

test('withSheetsAdded opens a self closing sheets element', () => {
  const out = withSheetsAdded('<workbook><sheets/></workbook>', [one])
  assert.match(out, /<sheets><sheet name="New"[^>]*\/><\/sheets>/)
})

test('withSheetsAdded keeps the workbook element prefix', () => {
  const out = withSheetsAdded('<x:workbook><x:sheets><x:sheet/></x:sheets></x:workbook>', [one])
  assert.match(out, /<x:sheet name="New" sheetId="5" r:id="rId9"\/><\/x:sheets>/)
})

test('withSheetsAdded refuses a workbook with no sheets element', () => {
  assert.throws(() => withSheetsAdded('<workbook></workbook>', [one]), isInvalid)
})

test('the wiring leaves a part alone when nothing is added', () => {
  assert.equal(
    withSheetsAdded('<workbook><sheets/></workbook>', []),
    '<workbook><sheets/></workbook>',
  )
  assert.equal(withSheetRelationships('<Relationships/>', []), '<Relationships/>')
  assert.equal(withSheetContentTypes('<Types/>', []), '<Types/>')
})

test('withSheetRelationships adds a worksheet relationship and refuses a malformed part', () => {
  assert.match(
    withSheetRelationships('<Relationships></Relationships>', [one]),
    /Id="rId9"[^>]*Target="worksheets\/sheet3.xml"/,
  )
  assert.throws(() => withSheetRelationships('<Relationships>', [one]), isInvalid)
})

test('withSheetContentTypes declares the sheet part and refuses a malformed part', () => {
  assert.match(
    withSheetContentTypes('<Types></Types>', [one]),
    /PartName="\/xl\/worksheets\/sheet3.xml"/,
  )
  assert.throws(() => withSheetContentTypes('<Types>', [one]), isInvalid)
})

test('withSheetsAdded escapes a name for the attribute', () => {
  const out = withSheetsAdded('<workbook><sheets/></workbook>', [
    added('A & B', 'xl/worksheets/s.xml'),
  ])
  assert.match(out, /name="A &amp; B"/)
})

test('withSheetRenamed rewrites the matching sheet, keeping the prefix', () => {
  assert.match(
    withSheetRenamed(
      '<workbook><sheets><sheet name="Old" sheetId="1" r:id="rId1"/></sheets></workbook>',
      'Old',
      'New',
    ),
    /<sheet name="New" sheetId="1"/,
  )
  assert.match(
    withSheetRenamed('<x:sheets><x:sheet name="Old" r:id="rId1"/></x:sheets>', 'Old', 'New'),
    /<x:sheet name="New"/,
  )
})

test('withSheetRenamed escapes the new name and leaves an unmatched name alone', () => {
  assert.match(
    withSheetRenamed('<sheets><sheet name="A"/></sheets>', 'A', 'X & Y'),
    /name="X &amp; Y"/,
  )
  const xml = '<sheets><sheet name="A"/></sheets>'
  assert.equal(withSheetRenamed(xml, 'Missing', 'New'), xml)
})

test('withSheetRemoved deletes the matching sheet, keeping the rest and the prefix', () => {
  assert.equal(
    withSheetRemoved(
      '<sheets><sheet name="A" r:id="rId1"/><sheet name="B" r:id="rId2"/></sheets>',
      'A',
    ),
    '<sheets><sheet name="B" r:id="rId2"/></sheets>',
  )
  assert.equal(
    withSheetRemoved('<x:sheets><x:sheet name="A"/></x:sheets>', 'A'),
    '<x:sheets></x:sheets>',
  )
  const xml = '<sheets><sheet name="A"/></sheets>'
  assert.equal(withSheetRemoved(xml, 'Missing'), xml)
})

test('checkSheetName accepts a valid name and refuses the rest', () => {
  checkSheetName('Fine', ['Other'])
  const refuses = (name: unknown, existing: readonly string[] = []) =>
    assert.throws(() => checkSheetName(name, existing), isUnwritable)
  refuses(123) // not a string
  refuses('') // empty
  refuses('a'.repeat(32)) // longer than 31
  refuses('a:b') // a character Excel forbids
  refuses('Same', ['same']) // a name already taken, in another case
})
