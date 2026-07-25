import assert from 'node:assert/strict'
import { test } from 'node:test'
import { checkDefinedName, readDefinedNames, withDefinedNames } from './defined-names.js'
import { XlsxError } from './errors.js'

const isUnwritable = (error: unknown) =>
  error instanceof XlsxError && error.code === 'unwritable-value'

test('readDefinedNames reads global names, skipping sheet-scoped ones', () => {
  const xml =
    '<workbook><definedNames><definedName name="Tax">Sheet1!$B$1</definedName>' +
    '<definedName name="Local" localSheetId="0">Sheet1!$A$1</definedName></definedNames></workbook>'
  assert.deepEqual([...readDefinedNames(xml)], [['Tax', 'Sheet1!$B$1']])
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
