import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  ensureAlignmentStyle,
  ensureBorderStyle,
  ensureDateStyle,
  ensureFillStyle,
  ensureFontStyle,
  readFormatting,
} from './styles-writer.js'
import { assertWellFormed } from '../testing/invariants.js'
import { ensureNumberFormat } from './styles-writer.js'
import { isDateFormat, numberFormatOf, readStyles } from './styles.js'

const xfAt = (xml: string, index: number) =>
  [...xml.matchAll(/<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/g)].map((match) => match[0])[index] ??
  ''

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

test('seeds a default font and lands ours past it when the file has none', () => {
  const source =
    '<styleSheet><cellXfs count="1">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, { bold: true })

  assert.match(result.xml, /<fonts count="2"><font\/><font><b\/><\/font><\/fonts><cellXfs/)
  assert.match(xfAt(result.xml, result.index), /fontId="1"/)
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

test('applying strikethrough adds a strike to the font', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFontStyle(source, 0, { strike: true })

  assert.match(result.xml, /<font><strike\/><sz val="11"\/><name val="Calibri"\/><\/font>/)
})

test('a double underline is written with its val', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  assert.match(ensureFontStyle(source, 0, { underline: 'double' }).xml, /<u val="double"\/>/)
})

test('true and single underline both write a bare u', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  assert.match(ensureFontStyle(source, 0, { underline: true }).xml, /<font><u\/>/)
  assert.match(ensureFontStyle(source, 0, { underline: 'single' }).xml, /<font><u\/>/)
})

test('a superscript is written as a vertAlign', () => {
  const source = fontStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  assert.match(
    ensureFontStyle(source, 0, { vertAlign: 'superscript' }).xml,
    /<vertAlign val="superscript"\/>/,
  )
})

test('strike, an underline style and vertAlign merge onto the cell font in order', () => {
  const source =
    '<styleSheet><fonts count="1"><font><b/><sz val="14"/></font></fonts>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs></styleSheet>'

  const result = ensureFontStyle(source, 0, {
    strike: true,
    underline: 'double',
    vertAlign: 'subscript',
  })

  assert.match(
    result.xml,
    /<font><b\/><strike\/><u val="double"\/><vertAlign val="subscript"\/><sz val="14"\/><\/font>/,
  )
})

const fills =
  '<fills count="2"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill></fills>'

const fillStyles = (cellXfs: string) =>
  `<styleSheet>${fills}<cellXfs count="1">${cellXfs}</cellXfs></styleSheet>`

test('applying a fill adds a solid fill and points the cell format at it', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFillStyle(source, 0, { type: 'solid', color: 'FFFF00' })

  assert.match(
    result.xml,
    /<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"\/><bgColor indexed="64"\/><\/patternFill><\/fill><\/fills>/,
  )
  assert.match(xfAt(result.xml, result.index), /fillId="2"/)
  assert.match(xfAt(result.xml, result.index), /applyFill="1"/)
})

test('a solid fill lands past the two reserved fill ids', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFillStyle(source, 0, { type: 'solid', color: '000000' })

  assert.match(result.xml, /<fills count="3">/)
  assert.match(xfAt(result.xml, result.index), /fillId="2"/)
})

test('a pattern fill carries its pattern, foreground and background', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFillStyle(source, 0, {
    type: 'pattern',
    pattern: 'lightGrid',
    color: 'FF0000',
    background: 'FFFFFF',
  })

  assert.match(
    result.xml,
    /<fill><patternFill patternType="lightGrid"><fgColor rgb="FFFF0000"\/><bgColor rgb="FFFFFFFF"\/><\/patternFill><\/fill>/,
  )
})

test('a pattern fill without a background falls back to the default indexed one', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureFillStyle(source, 0, {
    type: 'pattern',
    pattern: 'darkTrellis',
    color: '112233',
  })

  assert.match(
    result.xml,
    /<patternFill patternType="darkTrellis"><fgColor rgb="FF112233"\/><bgColor indexed="64"\/><\/patternFill>/,
  )
})

test('seeds the reserved fills when a file has no fills table', () => {
  const source =
    '<styleSheet><cellXfs count="1">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>'

  const result = ensureFillStyle(source, 0, { type: 'solid', color: '000000' })

  // the two reserved fills are seeded, then ours lands at id 2
  assert.match(
    result.xml,
    /<fills count="3"><fill><patternFill patternType="none"\/><\/fill><fill><patternFill patternType="gray125"\/><\/fill><fill><patternFill patternType="solid"/,
  )
  assert.match(xfAt(result.xml, result.index), /fillId="2"/)
})

test('does not add the same fill twice', () => {
  const source = fillStyles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const once = ensureFillStyle(source, 0, { type: 'solid', color: 'FF0000' })
  const twice = ensureFillStyle(once.xml, 0, { type: 'solid', color: 'FF0000' })

  assert.equal(once.index, twice.index)
  assert.equal(once.xml, twice.xml)
})

const bordersTable =
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'

const borderStyles = (cellXfs: string) =>
  `<styleSheet>${bordersTable}<cellXfs count="1">${cellXfs}</cellXfs></styleSheet>`

const emptyBorderXf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

test('applying a border to all sides adds one and points the cell format at it', () => {
  const result = ensureBorderStyle(borderStyles(emptyBorderXf), 0, { all: { style: 'thin' } })

  assert.match(
    result.xml,
    /<border><left style="thin"\/><right style="thin"\/><top style="thin"\/><bottom style="thin"\/><diagonal\/><\/border><\/borders>/,
  )
  assert.match(xfAt(result.xml, result.index), /borderId="1"/)
  assert.match(xfAt(result.xml, result.index), /applyBorder="1"/)
})

test('a border colour is written as a child of the side', () => {
  const result = ensureBorderStyle(borderStyles(emptyBorderXf), 0, {
    bottom: { style: 'medium', color: 'FF0000' },
  })

  assert.match(result.xml, /<bottom style="medium"><color rgb="FFFF0000"\/><\/bottom>/)
})

test('a border change merges onto the sides the cell already has', () => {
  const source =
    '<styleSheet><borders count="1"><border><left style="thin"/><right/><top/><bottom/>' +
    '<diagonal/></border></borders>' +
    `<cellXfs count="1">${emptyBorderXf}</cellXfs></styleSheet>`

  const result = ensureBorderStyle(source, 0, { top: { style: 'thick' } })

  assert.match(
    result.xml,
    /<border><left style="thin"\/><right\/><top style="thick"\/><bottom\/><diagonal\/><\/border>/,
  )
})

test('a border change keeps a coloured side the cell already has', () => {
  const source =
    '<styleSheet><borders count="1"><border>' +
    '<left style="thin"><color rgb="FFFF0000"/></left><right/><top/><bottom/><diagonal/>' +
    '</border></borders>' +
    `<cellXfs count="1">${emptyBorderXf}</cellXfs></styleSheet>`

  const result = ensureBorderStyle(source, 0, { top: { style: 'thick' } })

  assert.match(result.xml, /<left style="thin"><color rgb="FFFF0000"\/><\/left>/)
  assert.match(result.xml, /<top style="thick"\/>/)
})

test('ignores a border colour on a side that has no style', () => {
  const source =
    '<styleSheet><borders count="1"><border>' +
    '<left><color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/>' +
    '</border></borders>' +
    `<cellXfs count="1">${emptyBorderXf}</cellXfs></styleSheet>`

  const result = ensureBorderStyle(source, 0, { top: { style: 'thin' } })

  // the colour on a styleless side is dropped, not carried as a border
  assert.match(result.xml, /<border><left\/><right\/><top style="thin"\/>/)
})

test('a specific side overrides the all side', () => {
  const result = ensureBorderStyle(borderStyles(emptyBorderXf), 0, {
    all: { style: 'thin' },
    top: { style: 'thick' },
  })

  assert.match(result.xml, /<top style="thick"\/><bottom style="thin"\/>/)
})

test('seeds an empty border when the file has no borders table', () => {
  const source = `<styleSheet><cellXfs count="1">${emptyBorderXf}</cellXfs></styleSheet>`

  const result = ensureBorderStyle(source, 0, { all: { style: 'thin' } })

  assert.match(
    result.xml,
    /<borders count="2"><border><left\/><right\/><top\/><bottom\/><diagonal\/><\/border><border><left style="thin"/,
  )
  assert.match(xfAt(result.xml, result.index), /borderId="1"/)
})

test('does not add the same border twice', () => {
  const source = borderStyles(emptyBorderXf)

  const once = ensureBorderStyle(source, 0, { all: { style: 'thin' } })
  const twice = ensureBorderStyle(once.xml, 0, { all: { style: 'thin' } })

  assert.equal(once.index, twice.index)
  assert.equal(once.xml, twice.xml)
})

test('adds a prefixed font to a prefixed styles table', () => {
  const source =
    '<x:styleSheet><x:fonts count="1"><x:font/></x:fonts>' +
    '<x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>'

  const result = ensureFontStyle(source, 0, { bold: true, name: 'Courier' })

  assertWellFormed(result.xml, 'prefixed font')
  assert.match(result.xml, /<x:font><x:b\/><x:name val="Courier"\/><\/x:font>/)
})

test('adds a prefixed fill to a prefixed styles table', () => {
  const source =
    '<x:styleSheet><x:cellXfs count="1"><x:xf numFmtId="0" fillId="0"/></x:cellXfs></x:styleSheet>'

  const result = ensureFillStyle(source, 0, { type: 'solid', color: 'FF0000' })

  assertWellFormed(result.xml, 'prefixed fill')
  assert.match(
    result.xml,
    /<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FFFF0000"\/><x:bgColor indexed="64"\/><\/x:patternFill><\/x:fill>/,
  )
})

test('adds a prefixed border, colour and all, to a prefixed styles table', () => {
  const source =
    '<x:styleSheet><x:cellXfs count="1"><x:xf numFmtId="0" borderId="0"/></x:cellXfs></x:styleSheet>'

  const result = ensureBorderStyle(source, 0, { top: { style: 'thick', color: '112233' } })

  assertWellFormed(result.xml, 'prefixed border')
  assert.match(result.xml, /<x:top style="thick"><x:color rgb="FF112233"\/><\/x:top>/)
})

test("reads a cell format's font, fill and border", () => {
  const source =
    '<styleSheet>' +
    '<fonts count="2"><font/><font><b/><sz val="12"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FF000000"/></left><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1"/></cellXfs></styleSheet>'

  const formatting = readFormatting(source)

  assert.deepEqual(formatting[0], {})
  assert.deepEqual(formatting[1], {
    font: { bold: true, size: 12 },
    fill: { type: 'solid', color: 'FF00FF00' },
    border: { left: { style: 'thin', color: 'FF000000' } },
  })
})

test('reads a pattern fill with its foreground and background', () => {
  const source =
    '<styleSheet><fills count="3"><fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="lightGrid"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFFFF"/>' +
    '</patternFill></fill></fills>' +
    '<cellXfs count="1"><xf fillId="2"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], {
    fill: { type: 'pattern', pattern: 'lightGrid', color: 'FFFF0000', background: 'FFFFFFFF' },
  })
})

test('an unknown pattern type is not reported as a fill', () => {
  const source =
    '<styleSheet><fills count="1">' +
    '<fill><patternFill patternType="woven"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>' +
    '<cellXfs count="1"><xf fillId="0"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], {})
})

test('reports no font for a non-default id that points at an empty font', () => {
  const source =
    '<styleSheet><fonts count="2"><font/><font/></fonts>' +
    '<cellXfs count="1"><xf fontId="1"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], {})
})

test('ignores an unknown border style', () => {
  const source =
    '<styleSheet><borders count="1"><border><left style="weird"/><right/><top/><bottom/>' +
    '<diagonal/></border></borders><cellXfs count="1"><xf borderId="0"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], {})
})

test('reads a border side with a style but no colour, and skips a colourless fill', () => {
  const source =
    '<styleSheet><fills count="1"><fill><patternFill patternType="solid"/></fill></fills>' +
    '<borders count="1"><border><top style="thin"/><left/><right/><bottom/><diagonal/></border></borders>' +
    '<cellXfs count="1"><xf fillId="0" borderId="0"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], { border: { top: { style: 'thin' } } })
})

test('reports nothing for ids past the end of a table', () => {
  const source =
    '<styleSheet><fonts count="1"><font><b/></font></fonts>' +
    '<cellXfs count="1"><xf fontId="9" fillId="9" borderId="9"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], {})
})

test('reads nothing from a styles part with no tables', () => {
  assert.deepEqual(readFormatting('<styleSheet></styleSheet>'), [])
})

test("reads a cell format's alignment", () => {
  const source =
    '<styleSheet><cellXfs count="2"><xf/>' +
    '<xf applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"' +
    ' textRotation="90" indent="2"/></xf></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[1], {
    alignment: {
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      textRotation: 90,
      indent: 2,
    },
  })
})

test('ignores an unknown horizontal or vertical alignment', () => {
  const source =
    '<styleSheet><cellXfs count="1">' +
    '<xf><alignment horizontal="sideways" vertical="middle"/></xf></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[0], {})
})

test("reads a font's strike, underline style and vertAlign", () => {
  const source =
    '<styleSheet><fonts count="2"><font/>' +
    '<font><strike/><u val="double"/><vertAlign val="superscript"/></font></fonts>' +
    '<cellXfs count="2"><xf/><xf fontId="1"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[1], {
    font: { strike: true, underline: 'double', vertAlign: 'superscript' },
  })
})

test('a plain u reads back as a boolean underline', () => {
  const source =
    '<styleSheet><fonts count="2"><font/><font><u/></font></fonts>' +
    '<cellXfs count="2"><xf/><xf fontId="1"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[1], { font: { underline: true } })
})

test('an unknown underline val reads back as a boolean underline', () => {
  const source =
    '<styleSheet><fonts count="2"><font/><font><u val="wavy"/></font></fonts>' +
    '<cellXfs count="2"><xf/><xf fontId="1"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[1], { font: { underline: true } })
})

test('an explicit u val=none is not reported as underlined', () => {
  const source =
    '<styleSheet><fonts count="2"><font/><font><u val="none"/><b/></font></fonts>' +
    '<cellXfs count="2"><xf/><xf fontId="1"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[1], { font: { bold: true } })
})

test('an unknown vertAlign is ignored', () => {
  const source =
    '<styleSheet><fonts count="2"><font/><font><vertAlign val="sideways"/><b/></font></fonts>' +
    '<cellXfs count="2"><xf/><xf fontId="1"/></cellXfs></styleSheet>'

  assert.deepEqual(readFormatting(source)[1], { font: { bold: true } })
})

test('reports no alignment for a cell format that has none', () => {
  assert.deepEqual(
    readFormatting('<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>')[0],
    {},
  )
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

test('centering a cell adds an alignment and turns applyAlignment on', () => {
  const source = styles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

  const result = ensureAlignmentStyle(source, 0, { horizontal: 'center', vertical: 'top' })

  assert.equal(result.index, 1)
  assertWellFormed(result.xml, 'alignment styles')
  assert.match(xfAt(result.xml, 1), /applyAlignment="1"/)
  assert.match(xfAt(result.xml, 1), /<alignment horizontal="center" vertical="top"\/>/)
})

test('wrap, rotation and indent are written as their attributes', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureAlignmentStyle(source, 0, { wrapText: true, textRotation: 90, indent: 2 })

  assert.match(xfAt(result.xml, 1), /<alignment wrapText="1" textRotation="90" indent="2"\/>/)
})

test('an alignment change merges onto the one the cell already has', () => {
  const source = styles('<xf numFmtId="0"><alignment horizontal="center"/></xf>')

  const result = ensureAlignmentStyle(source, 0, { vertical: 'bottom' })

  assert.match(xfAt(result.xml, 1), /<alignment horizontal="center" vertical="bottom"\/>/)
})

test('does not add the same alignment twice', () => {
  const source = styles('<xf numFmtId="0"/>')
  const once = ensureAlignmentStyle(source, 0, { horizontal: 'right' })
  const twice = ensureAlignmentStyle(once.xml, 0, { horizontal: 'right' })

  assert.equal(twice.index, once.index)
  assert.equal(twice.xml, once.xml)
})

test('a full alignment on the base is kept when one attribute is changed', () => {
  const source = styles(
    '<xf numFmtId="0"><alignment horizontal="left" vertical="bottom" wrapText="1"' +
      ' textRotation="45" indent="3"/></xf>',
  )

  const result = ensureAlignmentStyle(source, 0, { horizontal: 'right' })

  assert.match(
    xfAt(result.xml, 1),
    /<alignment horizontal="right" vertical="bottom" wrapText="1" textRotation="45" indent="3"\/>/,
  )
})

test('accepts a text rotation of 255, which stacks the text', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureAlignmentStyle(source, 0, { textRotation: 255 })

  assert.match(xfAt(result.xml, 1), /textRotation="255"/)
})

test('ignores an unreadable rotation on the base alignment', () => {
  const source = styles('<xf numFmtId="0"><alignment textRotation="tilted"/></xf>')

  const result = ensureAlignmentStyle(source, 0, { horizontal: 'center' })

  assert.match(xfAt(result.xml, 1), /<alignment horizontal="center"\/>/)
})

test('falls back to the default when the base cell format is missing', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureAlignmentStyle(source, 99, { horizontal: 'center' })

  assert.match(xfAt(result.xml, 1), /<alignment horizontal="center"\/>/)
})

test('reopens a self closing cell format to hold the alignment', () => {
  const source = styles('<xf numFmtId="0" fontId="0"/>')

  const result = ensureAlignmentStyle(source, 0, { horizontal: 'left' })

  assertWellFormed(result.xml, 'reopened xf')
  assert.match(
    xfAt(result.xml, 1),
    /^<xf [^>]*applyAlignment="1"><alignment horizontal="left"\/><\/xf>$/,
  )
})

test('formats a cell that has no style yet against the default', () => {
  const source = styles('<xf numFmtId="0"/>')

  const result = ensureAlignmentStyle(source, undefined, { horizontal: 'center' })

  assert.equal(result.index, 1)
  assert.match(xfAt(result.xml, 1), /<alignment horizontal="center"\/>/)
})

test('writes the alignment into a prefixed cell format', () => {
  const source =
    '<x:styleSheet><x:cellXfs count="1"><x:xf numFmtId="0"/></x:cellXfs></x:styleSheet>'

  const result = ensureAlignmentStyle(source, 0, { horizontal: 'center' })

  assertWellFormed(result.xml, 'prefixed alignment')
  assert.match(
    result.xml,
    /<x:xf [^>]*applyAlignment="1"><x:alignment horizontal="center"\/><\/x:xf>/,
  )
})

test('refuses a text rotation outside the range a cell can hold', () => {
  const source = styles('<xf numFmtId="0"/>')

  assert.throws(() => ensureAlignmentStyle(source, 0, { textRotation: 300 }), /rotation/)
})

test('refuses a negative indent', () => {
  const source = styles('<xf numFmtId="0"/>')

  assert.throws(() => ensureAlignmentStyle(source, 0, { indent: -1 }), /[Ii]ndent/)
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
