import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createConditionalFormatStore, planConditionalFormat } from './conditional-format.js'
import { XlsxError } from './errors.js'
import type { ConditionalFormatSpec } from './patch.js'

const AT = { sheet: 'Sheet1', part: 'xl/worksheets/sheet1.xml' }
const STYLES = '<styleSheet></styleSheet>'

test('planConditionalFormat maps a colour scale without touching the styles table', () => {
  const result = planConditionalFormat(
    { colorScale: { min: 'FF0000', mid: '00FF00', max: '0000FF' } },
    'A1:A9',
    AT,
    STYLES,
  )
  assert.deepEqual(result.specs, [
    { kind: 'colorScale', sqref: 'A1:A9', colors: ['FFFF0000', 'FF00FF00', 'FF0000FF'] },
  ])
  assert.equal(result.dxfStyles, STYLES)
  assert.deepEqual(result.dxfColors, [])
})

test('planConditionalFormat maps a data bar', () => {
  const result = planConditionalFormat({ dataBar: { color: '638EC6' } }, 'A1:A9', AT, undefined)
  assert.deepEqual(result.specs, [{ kind: 'dataBar', sqref: 'A1:A9', color: 'FF638EC6' }])
})

test('planConditionalFormat allocates a dxf for a cellIs highlight and reports the colour', () => {
  const result = planConditionalFormat(
    { cellIs: { when: { greaterThan: 5 }, fill: 'FFFF00' } },
    'A1:A9',
    AT,
    STYLES,
  )
  assert.deepEqual(result.specs, [
    { kind: 'cellIs', sqref: 'A1:A9', operator: 'greaterThan', formulas: ['5'], dxfId: 0 },
  ])
  assert.deepEqual(result.dxfColors, ['FFFFFF00'])
  assert.notEqual(result.dxfStyles, STYLES)
})

test('planConditionalFormat refuses a highlight when the package has no style table', () => {
  assert.throws(
    () => planConditionalFormat({ duplicates: { fill: 'FF0000' } }, 'A1:A9', AT, undefined),
    (error: unknown) => error instanceof XlsxError && error.code === 'missing-part',
  )
})

test('planConditionalFormat refuses a non-finite comparison bound', () => {
  assert.throws(
    () =>
      planConditionalFormat(
        { cellIs: { when: { greaterThan: Number.POSITIVE_INFINITY }, fill: 'FF0000' } },
        'A1:A9',
        AT,
        STYLES,
      ),
    /not a finite number/,
  )
})

test('planConditionalFormat refuses a rank that is not a positive whole number', () => {
  assert.throws(
    () => planConditionalFormat({ top: { count: 0, fill: 'FF0000' } }, 'A1:A9', AT, STYLES),
    /positive whole number/,
  )
})

test('planConditionalFormat maps top, expression, duplicates and unique highlights', () => {
  const top = planConditionalFormat(
    { bottom: { count: 3, fill: 'FF0000', percent: true } },
    'A1:A9',
    AT,
    STYLES,
  )
  assert.deepEqual(top.specs, [
    { kind: 'top10', sqref: 'A1:A9', rank: 3, bottom: true, percent: true, dxfId: 0 },
  ])
  const expression = planConditionalFormat(
    { expression: { formula: 'A1>0', fill: 'FF0000' } },
    'A1:A9',
    AT,
    STYLES,
  )
  assert.equal(expression.specs[0]?.kind, 'expression')
  const unique = planConditionalFormat({ unique: { fill: 'FF0000' } }, 'A1:A9', AT, STYLES)
  assert.equal(unique.specs[0]?.kind, 'uniqueValues')
})

test('a data bar round-trips without any styles', () => {
  const store = createConditionalFormatStore()
  store.add('s', { kind: 'dataBar', sqref: 'A1:A9', color: 'FF638EC6' })
  assert.deepEqual(store.applied('s', [], undefined), [
    { range: 'A1:A9', rule: { dataBar: { color: 'FF638EC6' } } },
  ])
})

test('a two- and three-stop colour scale map back', () => {
  const store = createConditionalFormatStore()
  const two: ConditionalFormatSpec = {
    kind: 'colorScale',
    sqref: 'B1:B9',
    colors: ['FF0000', '00FF00'],
  }
  const three: ConditionalFormatSpec = {
    kind: 'colorScale',
    sqref: 'C1:C9',
    colors: ['FF0000', 'FFFF00', '00FF00'],
  }
  const applied = store.applied('s', [two, three], undefined)
  assert.deepEqual(applied[0]?.rule, { colorScale: { min: 'FF0000', max: '00FF00' } })
  assert.deepEqual(applied[1]?.rule, {
    colorScale: { min: 'FF0000', mid: 'FFFF00', max: '00FF00' },
  })
})

test('a colour scale short of two stops is left out', () => {
  const store = createConditionalFormatStore()
  const oneStop: ConditionalFormatSpec = { kind: 'colorScale', sqref: 'A1', colors: ['FF0000'] }
  assert.deepEqual(store.applied('s', [oneStop], undefined), [])
})

test('a dxf-backed rule with no style table to resolve it is left out', () => {
  const store = createConditionalFormatStore()
  const dup: ConditionalFormatSpec = { kind: 'duplicateValues', sqref: 'A1:A9', dxfId: 0 }
  assert.deepEqual(store.applied('s', [dup], undefined), [])
})

test('clear drops file rules but keeps one added afterward', () => {
  const store = createConditionalFormatStore()
  const fileBar: ConditionalFormatSpec = { kind: 'dataBar', sqref: 'Z9', color: 'FF000000' }
  store.clear('s')
  store.add('s', { kind: 'dataBar', sqref: 'A1', color: 'FFFFFFFF' })
  assert.ok(store.cleared.has('s'))
  assert.deepEqual(store.applied('s', [fileBar], undefined), [
    { range: 'A1', rule: { dataBar: { color: 'FFFFFFFF' } } },
  ])
})

test('applied lists file rules ahead of pending ones', () => {
  const store = createConditionalFormatStore()
  const fileBar: ConditionalFormatSpec = { kind: 'dataBar', sqref: 'Z9', color: 'FF000000' }
  store.add('s', { kind: 'dataBar', sqref: 'A1', color: 'FFFFFFFF' })
  const applied = store.applied('s', [fileBar], undefined)
  assert.equal(applied[0]?.range, 'Z9')
  assert.equal(applied[1]?.range, 'A1')
})
