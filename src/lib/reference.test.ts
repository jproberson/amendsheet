import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  columnToIndex,
  formatReference,
  indexToColumn,
  parseReference,
  parseWritableReference,
} from './reference.js'

test('reads a reference as one based row and column', () => {
  assert.deepEqual(parseReference('A1'), { row: 1, column: 1 })
  assert.deepEqual(parseReference('B3'), { row: 3, column: 2 })
})

test('reads columns past the first letter', () => {
  assert.equal(columnToIndex('Z'), 26)
  assert.equal(columnToIndex('AA'), 27)
  assert.equal(columnToIndex('AB'), 28)
  assert.equal(columnToIndex('BZ'), 78)
  assert.equal(columnToIndex('XFD'), 16384)
})

test('writes columns past the first letter', () => {
  assert.equal(indexToColumn(26), 'Z')
  assert.equal(indexToColumn(27), 'AA')
  assert.equal(indexToColumn(28), 'AB')
  assert.equal(indexToColumn(78), 'BZ')
  assert.equal(indexToColumn(16384), 'XFD')
})

test('round trips every column in the sheet', () => {
  for (let index = 1; index <= 16384; index++) {
    assert.equal(columnToIndex(indexToColumn(index)), index)
  }
})

test('reads the last cell of a sheet', () => {
  assert.deepEqual(parseReference('XFD1048576'), { row: 1048576, column: 16384 })
})

test('ignores absolute markers', () => {
  assert.deepEqual(parseReference('$B$7'), { row: 7, column: 2 })
})

test('accepts lower case letters', () => {
  assert.deepEqual(parseReference('c2'), { row: 2, column: 3 })
})

test('writes a reference', () => {
  assert.equal(formatReference({ row: 1, column: 1 }), 'A1')
  assert.equal(formatReference({ row: 1048576, column: 16384 }), 'XFD1048576')
})

test('rejects a reference with no row', () => {
  assert.throws(() => parseReference('A'), /"A" is not a cell reference/)
})

test('rejects a reference with no column', () => {
  assert.throws(() => parseReference('12'), /"12" is not a cell reference/)
})

test('rejects a reference with trailing characters', () => {
  assert.throws(() => parseReference('A1:B2'), /"A1:B2" is not a cell reference/)
})

test('rejects an empty reference', () => {
  assert.throws(() => parseReference(''), /"" is not a cell reference/)
})

test('rejects a column outside the sheet', () => {
  assert.throws(() => indexToColumn(0), /Column 0 is outside the sheet/)
  assert.throws(() => indexToColumn(16385), /Column 16385 is outside the sheet/)
})

test('refuses to write a row past the end of a sheet', () => {
  assert.throws(() => parseWritableReference('A1048577'), /outside the sheet/)
})

test('refuses to write a column past the end of a sheet', () => {
  assert.throws(() => parseWritableReference('XFE1'), /outside the sheet/)
  assert.throws(() => parseWritableReference('AAAAA1'), /outside the sheet/)
})

test('writes to the last cell of a sheet', () => {
  assert.deepEqual(parseWritableReference('XFD1048576'), { row: 1048576, column: 16384 })
})

test('refuses to write to row zero', () => {
  assert.throws(() => parseWritableReference('A0'), /outside the sheet/)
})

test('still reads a row zero reference, which real files contain', () => {
  assert.deepEqual(parseReference('A0'), { row: 0, column: 1 })
})
