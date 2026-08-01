import assert from 'node:assert/strict'
import { test } from 'node:test'
import { type SharedMasters, toCell } from './cell-read.js'
import type { RawCell } from './sheet.js'
import type { Styles } from './styles.js'

const address = { row: 1, column: 1 }
const plain: Styles = { numberFormats: new Map(), cellFormats: [] }
// A cell format whose number-format id is the built-in date code 14 (m/d/yyyy).
const dateStyled: Styles = { numberFormats: new Map(), cellFormats: [14] }

const rawNumber = (value: number, styleIndex?: number): RawCell => ({
  address,
  reference: 'A1',
  value: { kind: 'number', value },
  ...(styleIndex === undefined ? {} : { styleIndex }),
})

test('a number under a date format reads back as a date carrying its serial', () => {
  const cell = toCell(rawNumber(44_000, 0), dateStyled, {}, false)
  assert.equal(cell.value.kind, 'date')
  assert.ok(cell.value.kind === 'date' && cell.value.serial === 44_000)
})

test('a literal ISO date reads back as a date', () => {
  const raw: RawCell = { address, reference: 'A1', value: { kind: 'date', value: '2020-01-15' } }
  const cell = toCell(raw, plain, {}, false)
  assert.equal(cell.value.kind, 'date')
})

test('a literal date that does not parse falls back to text', () => {
  const raw: RawCell = { address, reference: 'A1', value: { kind: 'date', value: 'not-a-date' } }
  const cell = toCell(raw, plain, {}, false)
  assert.deepEqual(cell.value, { kind: 'text', value: 'not-a-date' })
})

test('a number with no date format stays a number', () => {
  const cell = toCell(rawNumber(44_000, 0), plain, {}, false)
  assert.deepEqual(cell.value, { kind: 'number', value: 44_000 })
})

test('a serial outside the date range stays the number it is', () => {
  const cell = toCell(rawNumber(-5, 0), dateStyled, {}, false)
  assert.deepEqual(cell.value, { kind: 'number', value: -5 })
})

test('a shared-formula dependent resolves its master from the map', () => {
  const masters: SharedMasters = new Map([['3', 'A1']])
  const raw: RawCell = {
    address,
    reference: 'B2',
    value: { kind: 'number', value: 0 },
    formula: '',
    sharedIndex: '3',
  }
  const cell = toCell(raw, plain, {}, false, masters)
  assert.deepEqual(cell.formula, { kind: 'shared', master: 'A1' })
})

test('a shared-formula dependent with no known master is shared without one', () => {
  const raw: RawCell = {
    address,
    reference: 'B2',
    value: { kind: 'number', value: 0 },
    formula: '',
    sharedIndex: '3',
  }
  const cell = toCell(raw, plain, {}, false, new Map())
  assert.deepEqual(cell.formula, { kind: 'shared' })
})

test('the reference is canonicalised, not the file spelling', () => {
  const raw: RawCell = { address: { row: 1, column: 1 }, reference: 'a1', value: { kind: 'empty' } }
  assert.equal(toCell(raw, plain, {}, false).reference, 'A1')
})
