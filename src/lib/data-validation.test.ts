import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createValidationStore } from './data-validation.js'
import { XlsxError } from './errors.js'
import type { DataValidationSpec } from './patch.js'

const at = { part: 'xl/worksheets/sheet1.xml', sheet: 'Sheet1' }

test('add then applied maps a list rule back to itself', () => {
  const store = createValidationStore(false)
  store.add('sheet1', 'A1:A3', at, { list: ['x', 'y'] })
  assert.deepEqual(store.applied('sheet1', []), [
    { range: 'A1:A3', rule: { allowBlank: true, list: ['x', 'y'] } },
  ])
})

test('a whole-number between rule round-trips through the spec', () => {
  const store = createValidationStore(false)
  store.add('s', 'B2', at, { whole: { between: [1, 10] } })
  const [only] = store.applied('s', [])
  assert.deepEqual(only?.rule, { allowBlank: true, whole: { between: [1, 10] } })
})

test('a date rule survives the serial round-trip', () => {
  const store = createValidationStore(false)
  const bound = new Date(Date.UTC(2020, 0, 15))
  store.add('s', 'C1', at, { date: { greaterThan: bound } })
  const [only] = store.applied('s', [])
  assert.ok(only && 'date' in only.rule)
  assert.deepEqual(only.rule, { allowBlank: true, date: { greaterThan: bound } })
})

test('applied merges file specs ahead of pending ones', () => {
  const store = createValidationStore(false)
  const fileSpec: DataValidationSpec = {
    type: 'custom',
    sqref: 'Z9',
    allowBlank: true,
    formula1: 'TRUE()',
  }
  store.add('s', 'A1', at, { list: ['only'] })
  const applied = store.applied('s', [fileSpec])
  assert.equal(applied.length, 2)
  assert.equal(applied[0]?.range, 'Z9')
  assert.equal(applied[1]?.range, 'A1')
})

test('clear drops the file specs but keeps a rule added afterward', () => {
  const store = createValidationStore(false)
  const fileSpec: DataValidationSpec = {
    type: 'custom',
    sqref: 'Z9',
    allowBlank: true,
    formula1: 'TRUE()',
  }
  store.clear('s')
  store.add('s', 'A1', at, { list: ['kept'] })
  assert.ok(store.cleared.has('s'))
  const applied = store.applied('s', [fileSpec])
  assert.deepEqual(applied, [{ range: 'A1', rule: { allowBlank: true, list: ['kept'] } }])
})

test('a spec kind the model does not represent is left out', () => {
  const store = createValidationStore(false)
  const timeSpec: DataValidationSpec = {
    type: 'time',
    sqref: 'A1',
    allowBlank: true,
    formula1: '0.5',
  }
  assert.deepEqual(store.applied('s', [timeSpec]), [])
})

test('a list value holding a comma is refused with unwritable-value', () => {
  const store = createValidationStore(false)
  assert.throws(
    () => store.add('s', 'A1', at, { list: ['a,b'] }),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})

test('an empty list is refused', () => {
  const store = createValidationStore(false)
  assert.throws(
    () => store.add('s', 'A1', at, { list: [] }),
    (error: unknown) => error instanceof XlsxError && error.code === 'unwritable-value',
  )
})
