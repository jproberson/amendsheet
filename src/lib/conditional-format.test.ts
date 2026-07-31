import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createConditionalFormatStore } from './conditional-format.js'
import type { ConditionalFormatSpec } from './patch.js'

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
