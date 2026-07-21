import assert from 'node:assert/strict'
import { test } from 'node:test'
import { ensureDateStyle } from './styles-writer.js'
import { assertWellFormed } from '../testing/invariants.js'
import { isDateFormat, readStyles } from './styles.js'

const styles = (cellXfs: string, extra = '') =>
  `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${extra}<cellXfs count="${(cellXfs.match(/<xf/g) ?? []).length}">${cellXfs}</cellXfs></styleSheet>`

test('reuses a cell format that already shows dates', () => {
  const source = styles('<xf numFmtId="0"/><xf numFmtId="14"/>')

  const result = ensureDateStyle(source, undefined)

  assert.equal(result.index, 1)
  assert.equal(result.xml, source)
})

test('appends a cell format when none shows dates', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureDateStyle(source, undefined)

  assert.equal(result.index, 1)
  assert.match(result.xml, /<xf numFmtId="14"[^/]*\/><\/cellXfs>/)
})

test('raises the cell format count when it appends', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureDateStyle(source, undefined)

  assert.match(result.xml, /<cellXfs count="2">/)
})

test('the appended format reads back as a date', () => {
  const result = ensureDateStyle(styles('<xf numFmtId="0"/>'), undefined)

  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
})

test('keeps the other formatting of the style it is based on', () => {
  const source = styles('<xf numFmtId="0" fontId="3" fillId="5" borderId="2" applyFont="1"/>')

  const result = ensureDateStyle(source, 0)

  assert.match(result.xml, /<xf numFmtId="14" fontId="3" fillId="5" borderId="2" applyFont="1"/)
})

test('does not append twice for the same base style', () => {
  const source = styles('<xf numFmtId="0" fontId="3"/>')

  const once = ensureDateStyle(source, 0)
  const twice = ensureDateStyle(once.xml, 0)

  assert.equal(once.index, twice.index)
  assert.equal(once.xml, twice.xml)
})

test('leaves a base style that already shows dates alone', () => {
  const source = styles('<xf numFmtId="0"/><xf numFmtId="14" fontId="7"/>')

  const result = ensureDateStyle(source, 1)

  assert.equal(result.index, 1)
  assert.equal(result.xml, source)
})

test('honours a custom date format on the base style', () => {
  const source = styles(
    '<xf numFmtId="164"/>',
    '<numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>',
  )

  const result = ensureDateStyle(source, 0)

  assert.equal(result.index, 0)
  assert.equal(result.xml, source)
})

test('writes a cell format table into a styles part that has none', () => {
  const source = '<styleSheet xmlns="http://x"><fonts count="1"><font/></fonts></styleSheet>'

  const result = ensureDateStyle(source, undefined)

  assert.equal(result.index, 0)
  assert.match(result.xml, /<cellXfs count="1"><xf numFmtId="14"[^/]*\/><\/cellXfs><\/styleSheet>/)
})

test('handles a self closing cell format table', () => {
  const source = '<styleSheet xmlns="http://x"><cellXfs count="0"/></styleSheet>'

  const result = ensureDateStyle(source, undefined)

  assert.equal(result.index, 0)
  assert.match(result.xml, /<cellXfs count="1"><xf numFmtId="14"[^/]*\/><\/cellXfs>/)
})

test('adds a number format to a style that has none', () => {
  const source = styles('<xf fontId="2"/>')

  const result = ensureDateStyle(source, 0)

  assert.match(result.xml, /<xf numFmtId="14" fontId="2" applyNumberFormat="1"\/>/)
})

test('turns on number formatting that was switched off', () => {
  const source = styles('<xf numFmtId="0" fontId="2" applyNumberFormat="0"/>')

  const result = ensureDateStyle(source, 0)

  assert.match(result.xml, /<xf numFmtId="14" fontId="2" applyNumberFormat="1"\/>/)
})

test('falls back to a plain date format when the base style is missing', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureDateStyle(source, 99)

  assert.equal(result.index, 1)
  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
})

test('keeps a style written with a closing tag rather than self closed', () => {
  const source = styles('<xf numFmtId="0" fontId="1"><alignment horizontal="center"/></xf>')

  const result = ensureDateStyle(source, 0)

  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
})

test('clones a style that has children without breaking the table', () => {
  const source = styles(
    '<xf numFmtId="3" fontId="6" applyNumberFormat="1" applyAlignment="1">' +
      '<alignment horizontal="center"/></xf>',
  )

  const result = ensureDateStyle(source, 0)

  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
  assertWellFormed(result.xml, 'styles with children')
  assert.match(result.xml, /<alignment horizontal="center"\/><\/xf><\/cellXfs>/)
})
