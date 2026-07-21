import assert from 'node:assert/strict'
import { test } from 'node:test'
import { ensureDateStyle } from './styles-writer.js'
import { assertWellFormed } from '../testing/invariants.js'
import { ensureNumberFormat } from './styles-writer.js'
import { isDateFormat, numberFormatOf, readStyles } from './styles.js'

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

test('writes into a prefixed style table without breaking it', () => {
  const source =
    '<x:styleSheet><x:cellXfs count="1"><x:xf numFmtId="0"/></x:cellXfs></x:styleSheet>'

  const result = ensureDateStyle(source, undefined)

  assertWellFormed(result.xml, 'prefixed cellXfs')
  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
})

test('opens a prefixed self closing style table correctly', () => {
  const source = '<x:styleSheet><x:cellXfs count="0"/></x:styleSheet>'

  const result = ensureDateStyle(source, undefined)

  assertWellFormed(result.xml, 'prefixed self closing cellXfs')
  assert.match(result.xml, /<\/x:cellXfs>/)
})

test('adds a style table to a prefixed document that has none', () => {
  const source = '<x:styleSheet><x:fonts count="1"><x:font/></x:fonts></x:styleSheet>'

  const result = ensureDateStyle(source, undefined)

  assertWellFormed(result.xml, 'prefixed root with no cellXfs')
  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
})

test('turns on number formatting written as a word', () => {
  const source = styles('<xf numFmtId="0" fontId="2" applyNumberFormat="false"/>')

  const result = ensureDateStyle(source, 0)

  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
  assert.match(result.xml, /applyNumberFormat="1"/)
})

test('reuses a built in format when the code matches one', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureNumberFormat(source, undefined, '0.00%')

  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '0.00%')
  assert.equal(result.xml.includes('<numFmt'), false, 'a custom format was written needlessly')
})

test('writes a custom format the file does not have', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureNumberFormat(source, undefined, '"$"#,##0.00')

  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '"$"#,##0.00')
  assert.match(result.xml, /<numFmts[^>]*>/)
})

test('reuses a custom format the file already declares', () => {
  const source = styles(
    '<xf numFmtId="164"/>',
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>',
  )

  const result = ensureNumberFormat(source, undefined, '0.000')

  assert.equal(result.index, 0)
  assert.equal(result.xml, source)
})

test('picks an id that is not already taken, including inside dxfs', () => {
  const source = styles(
    '<xf numFmtId="0"/>',
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>' +
      '<dxfs><dxf><numFmt numFmtId="165" formatCode="0.0"/></dxf></dxfs>',
  )

  const result = ensureNumberFormat(source, undefined, '#,##0')

  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '#,##0')
  assert.equal(result.xml.includes('numFmtId="165" formatCode="#,##0"'), false)
})

test('keeps the other formatting of the style it is based on', () => {
  const source = styles('<xf numFmtId="0" fontId="3" fillId="5"/>')

  const result = ensureNumberFormat(source, 0, '0.0%')

  assert.match(result.xml, /fontId="3" fillId="5"/)
})

test('refuses a format code xml cannot hold', () => {
  assert.throws(
    () =>
      ensureNumberFormat(styles('<xf numFmtId="0"/>'), undefined, `0.0${String.fromCharCode(7)}`),
    /cannot be written to xml/i,
  )
})

test('adds a format table to a document that has none', () => {
  const source = '<styleSheet><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

  const result = ensureNumberFormat(source, undefined, '0.000')

  assert.match(
    result.xml,
    /<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"\/><\/numFmts>/,
  )
  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '0.000')
})

test('writes into a format table that was self closing', () => {
  const source =
    '<styleSheet><numFmts count="0"/><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

  const result = ensureNumberFormat(source, undefined, '0.000')

  assertWellFormed(result.xml, 'self closing numFmts')
  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '0.000')
})

test('writes a format into a prefixed document', () => {
  const source =
    '<x:styleSheet><x:cellXfs count="1"><x:xf numFmtId="0"/></x:cellXfs></x:styleSheet>'

  const result = ensureNumberFormat(source, undefined, '0.000')

  assertWellFormed(result.xml, 'prefixed numFmts')
  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '0.000')
})

test('refuses a document with no styleSheet element', () => {
  assert.throws(() => ensureNumberFormat('<other/>', undefined, '0.000'), /styleSheet/)
})
