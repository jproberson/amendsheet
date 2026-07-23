import assert from 'node:assert/strict'
import { test } from 'node:test'
import { ensureDateStyle, ensureFillStyle, ensureFontStyle } from './styles-writer.js'
import { assertWellFormed } from '../testing/invariants.js'
import { ensureNumberFormat } from './styles-writer.js'
import { isDateFormat, numberFormatOf, readStyles } from './styles.js'

const xfAt = (xml: string, index: number) =>
  [...xml.matchAll(/<xf [^>]*(?:\/>|>[\s\S]*?<\/xf>)/g)].map((match) => match[0])[index] ?? ''

const fontsTable = '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'

const fontStyles = (cellXfs: string) =>
  `<styleSheet>${fontsTable}<cellXfs count="${(cellXfs.match(/<xf/g) ?? []).length}">${cellXfs}</cellXfs></styleSheet>`

test('applying bold adds a font and a cell format that uses it', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFontStyle(source, 0, { bold: true })

  assert.match(result.xml, /<fonts count="2">/)
  assert.match(result.xml, /<font><b\/><sz val="11"\/><name val="Calibri"\/><\/font><\/fonts>/)
  assert.match(xfAt(result.xml, result.index), /fontId="1"/)
  assert.match(xfAt(result.xml, result.index), /applyFont="1"/)
})

test('a font change merges onto the cell font rather than replacing it', () => {
  const source =
    '<styleSheet><fonts count="1"><font><b/><sz val="14"/><color rgb="FFFF0000"/>' +
    '<name val="Arial"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, { italic: true })

  // bold, size, colour and name all survive; italic is added.
  assert.match(
    result.xml,
    /<font><b\/><i\/><sz val="14"\/><color rgb="FFFF0000"\/><name val="Arial"\/><\/font>/,
  )
})

test('a six digit colour is stored as opaque argb', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFontStyle(source, 0, { color: '00FF00' })

  assert.match(result.xml, /<color rgb="FF00FF00"\/>/)
})

test('the same font is not added twice', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const once = ensureFontStyle(source, 0, { bold: true })
  const twice = ensureFontStyle(once.xml, 0, { bold: true })

  assert.equal(once.index, twice.index)
  assert.equal(once.xml, twice.xml)
})

test('rejects a colour that is not hex', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  assert.throws(() => ensureFontStyle(source, 0, { color: 'red' }), /hex/)
})

test('accepts a colour written with a leading hash', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFontStyle(source, 0, { color: '#123456' })

  assert.match(result.xml, /<color rgb="FF123456"\/>/)
})

test('refuses to add a font to a document with no styleSheet', () => {
  const source = '<other><cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs></other>'

  assert.throws(() => ensureFontStyle(source, 0, { bold: true }), /styleSheet/)
})

test('a bold flag explicitly off is read as not bold', () => {
  const source =
    '<styleSheet><fonts count="1"><font><b val="0"/><sz val="11"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, { italic: true })

  assert.match(result.xml, /<font><i\/><sz val="11"\/><\/font>/)
})

test('formats a cell that has no style yet against the default font', () => {
  const source = fontStyles('')

  const result = ensureFontStyle(source, undefined, { underline: true, name: 'Courier' })

  assert.match(result.xml, /<font><u\/><sz val="11"\/><name val="Courier"\/><\/font>/)
})

test('creates a fonts table when the file has none', () => {
  const source =
    '<styleSheet><cellXfs count="1">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, { bold: true })

  assert.match(result.xml, /<fonts count="1"><font><b\/><\/font><\/fonts><cellXfs/)
})

test('ignores a themed colour and an unparseable size in the base font', () => {
  const source =
    '<styleSheet><fonts count="1"><font><sz/><color theme="1"/><name val="Calibri"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, { bold: true })

  assert.match(result.xml, /<font><b\/><name val="Calibri"\/><\/font>/)
})

test('treats a cell format with no fontId as using the default font', () => {
  const source =
    '<styleSheet><fonts count="1"><font><sz val="11"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, { bold: true })

  assert.match(result.xml, /<font><b\/><sz val="11"\/><\/font>/)
})

const fills =
  '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill></fills>'

const fillStyles = (cellXfs: string) =>
  `<styleSheet>${fills}<cellXfs count="1">${cellXfs}</cellXfs></styleSheet>`

test('applying a fill adds a solid fill and points the cell format at it', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFillStyle(source, 0, { color: 'FFFF00' })

  assert.match(
    result.xml,
    /<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"\/><bgColor indexed="64"\/><\/patternFill><\/fill><\/fills>/,
  )
  assert.match(xfAt(result.xml, result.index), /fillId="2"/)
  assert.match(xfAt(result.xml, result.index), /applyFill="1"/)
})

test('a solid fill lands past the two reserved fill ids', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFillStyle(source, 0, { color: '000000' })

  assert.match(result.xml, /<fills count="3">/)
  assert.match(xfAt(result.xml, result.index), /fillId="2"/)
})

test('seeds the reserved fills when a file has no fills table', () => {
  const source =
    '<styleSheet><cellXfs count="1">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFillStyle(source, 0, { color: '000000' })

  // the two reserved fills are seeded, then ours lands at id 2
  assert.match(
    result.xml,
    /<fills count="3"><fill><patternFill patternType="none"\/><\/fill><fill><patternFill patternType="gray125"\/><\/fill><fill><patternFill patternType="solid"/,
  )
  assert.match(xfAt(result.xml, result.index), /fillId="2"/)
})

test('does not add the same fill twice', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const once = ensureFillStyle(source, 0, { color: 'FF0000' })
  const twice = ensureFillStyle(once.xml, 0, { color: 'FF0000' })

  assert.equal(once.index, twice.index)
  assert.equal(once.xml, twice.xml)
})

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

test('keeps the font and fill of the cell being formatted', () => {
  // Index 1 already shows 0.00, but it belongs to a different cell. Borrowing
  // it would silently strip the bold and the fill from the cell being written.
  const source =
    '<styleSheet><cellXfs count="2">' +
    '<xf numFmtId="0" fontId="7" fillId="9" borderId="3" xfId="0"/>' +
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '</cellXfs></styleSheet>'

  const result = ensureNumberFormat(source, 0, '0.00')

  assertWellFormed(result.xml, 'kept formatting')
  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '0.00')
  const applied = [...result.xml.matchAll(/<xf [^>]*\/>/g)].map((match) => match[0])[result.index]
  assert.match(applied ?? '', /fontId="7"/)
  assert.match(applied ?? '', /fillId="9"/)
  assert.match(applied ?? '', /borderId="3"/)
})

test('counts the formats it appends to an existing table', () => {
  const source =
    '<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy"/></numFmts>' +
    '<cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

  const result = ensureNumberFormat(source, undefined, '"$"#,##0.00_);[Red]("$"#,##0.00)')

  const declared = /<numFmts count="(\d+)"/.exec(result.xml)?.[1]
  assert.equal(declared, '2', 'numFmts count must match its children or Excel offers to repair')
})

test('adds a count to a format table that left it off', () => {
  const source =
    '<styleSheet><numFmts><numFmt numFmtId="164" formatCode="yyyy"/></numFmts>' +
    '<cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

  const result = ensureNumberFormat(source, undefined, '0.000')

  assertWellFormed(result.xml, 'counted numFmts')
  assert.match(result.xml, /<numFmts count="2">/)
})

test('builds a format table when a stray close tag is all there is', () => {
  const source =
    '<styleSheet></numFmts><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>'

  const result = ensureNumberFormat(source, undefined, '0.000')

  assert.equal(numberFormatOf(readStyles(result.xml), result.index), '0.000')
  assert.match(result.xml, /<numFmts count="1">/)
})

test('applies a format to a prefixed cell format that declares no numFmtId', () => {
  // /^<xf/ cannot match <x:xf, so the format was dropped and the cell fell back
  // to General while still being written as a serial.
  const source =
    '<x:styleSheet><x:cellXfs count="1">' +
    '<x:xf fontId="1" fillId="0" borderId="0" xfId="0"/>' +
    '</x:cellXfs></x:styleSheet>'

  const result = ensureDateStyle(source, 0)

  assertWellFormed(result.xml, 'prefixed xf without numFmtId')
  assert.equal(
    isDateFormat(readStyles(result.xml), result.index),
    true,
    'a date needs a cell format that shows dates',
  )
  assert.match(result.xml, /fontId="1"/)
})

test('applies a format to a cell format whose attributes are single quoted', () => {
  // Either quote is legal xml, and matching only double quotes added a second
  // numFmtId rather than rewriting the one that was there.
  const source =
    "<styleSheet><cellXfs count='1'><xf numFmtId='0' fontId='3'/></cellXfs></styleSheet>"

  const result = ensureDateStyle(source, 0)

  assertWellFormed(result.xml, 'single quoted xf')
  assert.equal(isDateFormat(readStyles(result.xml), result.index), true)
  assert.match(result.xml, /fontId='3'/)
})

test('raises a single quoted cell format count', () => {
  const source = "<styleSheet><cellXfs count='1'><xf numFmtId='0'/></cellXfs></styleSheet>"

  const result = ensureDateStyle(source, undefined)

  assert.equal(readStyles(result.xml).cellFormats.length, 2)
  assert.doesNotMatch(result.xml, /count='1'/, 'the count still claims one format')
})

test('raises a single quoted number format count', () => {
  const source =
    "<styleSheet><numFmts count='1'><numFmt numFmtId='164' formatCode='yyyy'/></numFmts>" +
    "<cellXfs count='1'><xf numFmtId='0'/></cellXfs></styleSheet>"

  const result = ensureNumberFormat(source, undefined, '0.000')

  assertWellFormed(result.xml, 'single quoted numFmts')
  assert.doesNotMatch(result.xml, /numFmts count='1'/, 'the count still claims one format')
})
