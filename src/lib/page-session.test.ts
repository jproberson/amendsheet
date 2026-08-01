import assert from 'node:assert/strict'
import { test } from 'node:test'
import { XlsxError } from './errors.js'
import { createPageStore } from './page-session.js'

const at = { part: 'xl/worksheets/sheet1.xml', sheet: 'Sheet1' }
const isCode = (code: string) => (error: unknown) =>
  error instanceof XlsxError && error.code === code

test('page setup merges over the file and later edits merge over earlier ones', () => {
  const page = createPageStore()
  page.setPageSetup('s', { orientation: 'landscape' }, at)
  page.setPageSetup('s', { scale: 80 }, at)
  assert.deepEqual(page.mergedSetup('s', undefined), { orientation: 'landscape', scale: 80 })
})

test('a bad orientation and an out-of-range scale are refused', () => {
  const page = createPageStore()
  assert.throws(
    () => page.setPageSetup('s', { orientation: 'sideways' as 'portrait' }, at),
    isCode('unwritable-value'),
  )
  assert.throws(() => page.setPageSetup('s', { scale: 5 }, at), isCode('unwritable-value'))
})

test('a negative margin is refused', () => {
  const page = createPageStore()
  assert.throws(() => page.setPageMargins('s', { left: -1 }, at), isCode('unwritable-value'))
})

test('a row break stores one-based and reads back one-based, refusing row 1', () => {
  const page = createPageStore()
  assert.throws(() => page.addRowBreak('s', 1, at), isCode('bad-reference'))
  page.addRowBreak('s', 10, at)
  assert.deepEqual(page.mergedBreaks('s', undefined), { rows: [10], columns: [] })
})

test('a column break refuses column A and reads back its letter', () => {
  const page = createPageStore()
  assert.throws(() => page.addColumnBreak('s', 'A', at), isCode('bad-reference'))
  page.addColumnBreak('s', 'D', at)
  assert.deepEqual(page.mergedBreaks('s', undefined), { rows: [], columns: ['D'] })
})

test('hasPending and paths reflect what was set', () => {
  const page = createPageStore()
  assert.equal(page.hasPending(), false)
  assert.deepEqual([...page.paths()], [])
  page.setHeaderFooter('s', { header: { center: 'Title' } })
  assert.equal(page.hasPending(), true)
  assert.deepEqual([...page.paths()], ['s'])
})

test('apply leaves a sheet untouched when nothing is pending for its path', () => {
  const page = createPageStore()
  page.setPageSetup('other', { scale: 50 }, at)
  const xml = '<worksheet><sheetData/></worksheet>'
  assert.equal(page.apply(xml, 's'), xml)
})
