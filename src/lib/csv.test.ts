import assert from 'node:assert/strict'
import { test } from 'node:test'
import { formatCsv, parseCsv } from './csv.js'

test('parseCsv reads plain rows across CRLF, LF and a trailing newline', () => {
  assert.deepEqual(parseCsv('a,b,c\r\nd,e,f'), [
    ['a', 'b', 'c'],
    ['d', 'e', 'f'],
  ])
  assert.deepEqual(parseCsv('a,b\nc,d\n'), [
    ['a', 'b'],
    ['c', 'd'],
  ])
  assert.deepEqual(parseCsv(''), [])
})

test('parseCsv handles quotes, embedded delimiters, newlines and doubled quotes', () => {
  assert.deepEqual(parseCsv('"a,b","c""d","e\nf"'), [['a,b', 'c"d', 'e\nf']])
  assert.deepEqual(parseCsv('a,,c'), [['a', '', 'c']]) // empty field
  assert.deepEqual(parseCsv('a,\n'), [['a', '']]) // trailing empty field
  assert.deepEqual(parseCsv('a\n\nb'), [['a'], [''], ['b']]) // a blank line is one empty field
})

test('parseCsv takes a custom delimiter', () => {
  assert.deepEqual(parseCsv('a\tb\tc', '\t'), [['a', 'b', 'c']])
})

test('formatCsv quotes only fields that need it and doubles inner quotes', () => {
  assert.equal(
    formatCsv([
      ['a', 'b,c', 'd"e'],
      ['f\ng', 'h', 'i'],
    ]),
    'a,"b,c","d""e"\r\n"f\ng",h,i',
  )
})

test('a round trip through format and parse preserves the rows', () => {
  const rows = [
    ['name', 'note'],
    ['apple', 'red, round'],
    ['pear', 'has a "stem"'],
    ['multi', 'line\nhere'],
    ['', 'empty first'],
  ]
  assert.deepEqual(parseCsv(formatCsv(rows)), rows)
})
