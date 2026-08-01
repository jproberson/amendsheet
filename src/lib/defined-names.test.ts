import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  checkDefinedName,
  parsePrintTitles,
  printAreaRanges,
  printAreaRefersTo,
  printTitlesRefersTo,
  readDefinedNames,
  readSheetScopedNames,
  withDefinedNames,
} from './defined-names.js'
import { XlsxError } from './errors.js'

const isUnwritable = (error: unknown) =>
  error instanceof XlsxError && error.code === 'unwritable-value'

test('readDefinedNames reads global names, skipping sheet-scoped ones', () => {
  const xml =
    '<workbook><definedNames><definedName name="Tax">Sheet1!$B$1</definedName>' +
    '<definedName name="Local" localSheetId="0">Sheet1!$A$1</definedName></definedNames></workbook>'
  assert.deepEqual([...readDefinedNames(xml)], [['Tax', 'Sheet1!$B$1']])
})

test('readSheetScopedNames reads names carrying a localSheetId, skipping global ones', () => {
  const xml =
    '<workbook><definedNames><definedName name="Tax">Sheet1!$B$1</definedName>' +
    '<definedName name="Local" localSheetId="2">Sheet3!$A$1</definedName>' +
    '<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2</definedName>' +
    '</definedNames></workbook>'
  assert.deepEqual(readSheetScopedNames(xml), [
    { name: 'Local', localSheetId: 2, refersTo: 'Sheet3!$A$1' },
    { name: '_xlnm.Print_Area', localSheetId: 0, refersTo: 'Sheet1!$A$1:$B$2' },
  ])
})

test('withDefinedNames opens a container after sheets when there is none', () => {
  const xml = '<workbook><sheets><sheet name="S"/></sheets></workbook>'
  assert.match(
    withDefinedNames(xml, new Map([['A', 'S!$A$1']])),
    /<\/sheets><definedNames><definedName name="A">S!\$A\$1<\/definedName><\/definedNames>/,
  )
})

test('withDefinedNames inserts into an existing container and replaces redefined names', () => {
  const into =
    '<workbook><sheets/><definedNames><definedName name="A">a</definedName>' +
    '<definedName name="B">b</definedName></definedNames></workbook>'
  assert.match(
    withDefinedNames(into, new Map([['New', 'n']])),
    /<definedName name="B">b<\/definedName><definedName name="New">n<\/definedName>/,
  )
  // Redefining two at once removes both old spans, descending, and re-adds them once each.
  const out = withDefinedNames(
    into,
    new Map([
      ['A', 'x'],
      ['B', 'y'],
    ]),
  )
  assert.equal((out.match(/name="A"/g) ?? []).length, 1)
  assert.equal((out.match(/name="B"/g) ?? []).length, 1)
  assert.match(out, /name="A">x<\/definedName>/)
})

test('withDefinedNames keeps the prefix and refuses a workbook with no sheets', () => {
  assert.match(
    withDefinedNames('<x:workbook><x:sheets></x:sheets></x:workbook>', new Map([['A', 'v']])),
    /<x:definedNames><x:definedName name="A">v<\/x:definedName><\/x:definedNames>/,
  )
  assert.throws(
    () => withDefinedNames('<workbook></workbook>', new Map([['A', 'v']])),
    (error: unknown) => error instanceof XlsxError && error.code === 'invalid-content',
  )
})

test('withDefinedNames opens a self closing container and does nothing for no names', () => {
  assert.match(
    withDefinedNames('<workbook><sheets/><definedNames/></workbook>', new Map([['A', 'v']])),
    /<definedNames><definedName name="A">v<\/definedName><\/definedNames>/,
  )
  assert.equal(
    withDefinedNames('<workbook><sheets/></workbook>', new Map()),
    '<workbook><sheets/></workbook>',
  )
})

test('withDefinedNames writes a scoped name and replaces one matched by its pair', () => {
  const into =
    '<workbook><sheets/><definedNames>' +
    '<definedName name="R" localSheetId="0">old</definedName>' +
    '<definedName name="R" localSheetId="1">other sheet, kept</definedName>' +
    '</definedNames></workbook>'
  const out = withDefinedNames(into, new Map(), new Set(), {
    write: [{ name: 'R', localSheetId: 0, refersTo: 'new' }],
  })
  // Only the localSheetId=0 R is replaced; the localSheetId=1 R is untouched.
  assert.match(out, /<definedName name="R" localSheetId="1">other sheet, kept<\/definedName>/)
  assert.match(out, /<definedName name="R" localSheetId="0">new<\/definedName>/)
  assert.equal((out.match(/name="R" localSheetId="0"/g) ?? []).length, 1)
})

test('withDefinedNames removes a scoped name by its pair', () => {
  const into =
    '<workbook><sheets/><definedNames>' +
    '<definedName name="R" localSheetId="0">gone</definedName>' +
    '<definedName name="Keep" localSheetId="0">stays</definedName>' +
    '</definedNames></workbook>'
  const out = withDefinedNames(into, new Map(), new Set(), {
    remove: [{ name: 'R', localSheetId: 0 }],
  })
  assert.doesNotMatch(out, /name="R"/)
  assert.match(out, /<definedName name="Keep" localSheetId="0">stays<\/definedName>/)
})

test('printAreaRefersTo quotes the sheet and makes the range absolute', () => {
  assert.equal(printAreaRefersTo('My Sheet', 'A1:B2'), "'My Sheet'!$A$1:$B$2")
  assert.equal(printAreaRefersTo("O'Brien", 'A1:A1'), "'O''Brien'!$A$1:$A$1")
})

test('printAreaRanges strips the qualifier and $ across comma-joined ranges', () => {
  assert.equal(printAreaRanges("'S'!$A$1:$B$2,'S'!$D$1:$E$2"), 'A1:B2,D1:E2')
  assert.equal(printAreaRanges('$A$1:$C$3'), 'A1:C3')
})

test('print titles round-trip through refersTo and back, undefined when neither axis reads', () => {
  assert.equal(
    printTitlesRefersTo('Data', { rows: '1:2', columns: 'A:A' }),
    "'Data'!$A:$A,'Data'!$1:$2",
  )
  assert.deepEqual(parsePrintTitles("'Data'!$A:$A,'Data'!$1:$2"), { rows: '1:2', columns: 'A:A' })
  assert.equal(parsePrintTitles("'Data'!garbage"), undefined)
})

test('checkDefinedName accepts a valid name and refuses the rest', () => {
  checkDefinedName('Good_Name.1', 'Sheet1!$A$1')
  const refuses = (name: unknown, refersTo: unknown = 'x') =>
    assert.throws(() => checkDefinedName(name, refersTo), isUnwritable)
  refuses(123) // not a string
  refuses('') // empty
  refuses('1Bad') // a leading digit
  refuses('has space') // a space
  refuses('a'.repeat(256)) // longer than 255
  refuses('Good', '') // empty refersTo
  refuses('Good', 5) // non-string refersTo
})

test('withDefinedNames removes a name without re-adding it, keeping the rest', () => {
  const xml =
    '<workbook><sheets><sheet name="S"/></sheets><definedNames>' +
    '<definedName name="Tax">S!$A$1</definedName>' +
    '<definedName name="Keep">S!$B$1</definedName></definedNames></workbook>'
  const out = withDefinedNames(xml, new Map(), new Set(['Tax']))
  assert.doesNotMatch(out, /name="Tax"/)
  assert.match(out, /<definedName name="Keep">S!\$B\$1<\/definedName>/)
})
