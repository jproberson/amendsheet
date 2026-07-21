import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readContainer } from './container.js'
import { isDateFormat, numberFormatOf, readStyles } from './styles.js'

const styles = (body: string) =>
  `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</styleSheet>`

const withFormats = (codes: Array<[number, string]>, cellFormats: number[]) =>
  readStyles(
    styles(
      `<numFmts>${codes.map(([id, code]) => `<numFmt numFmtId="${id}" formatCode="${code}"/>`).join('')}</numFmts>` +
        `<cellXfs count="${cellFormats.length}">${cellFormats.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0"/>`).join('')}</cellXfs>`,
    ),
  )

test('reads custom number formats by id', () => {
  const parsed = withFormats([[164, 'yyyy-mm-dd']], [164])

  assert.equal(parsed.numberFormats.get(164), 'yyyy-mm-dd')
})

test('reads cell formats in index order', () => {
  const parsed = withFormats([], [0, 14, 164])

  assert.deepEqual(parsed.cellFormats, [0, 14, 164])
})

test('resolves the format code a cell uses', () => {
  const parsed = withFormats([[164, 'yyyy-mm-dd']], [0, 164])

  assert.equal(numberFormatOf(parsed, 1), 'yyyy-mm-dd')
})

test('resolves a built in format code', () => {
  const parsed = withFormats([], [14])

  assert.equal(numberFormatOf(parsed, 0), 'mm-dd-yy')
})

test('treats the built in date formats as dates', () => {
  for (const id of [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]) {
    const parsed = withFormats([], [id])
    assert.equal(isDateFormat(parsed, 0), true, `format ${id} should be a date`)
  }
})

test('treats the built in number formats as numbers', () => {
  for (const id of [0, 1, 2, 3, 4, 9, 10, 11, 37, 38, 39, 40, 48, 49]) {
    const parsed = withFormats([], [id])
    assert.equal(isDateFormat(parsed, 0), false, `format ${id} should not be a date`)
  }
})

test('treats a custom format with date tokens as a date', () => {
  const parsed = withFormats(
    [
      [164, 'yyyy-mm-dd'],
      [165, 'd mmm yyyy'],
      [166, 'hh:mm:ss'],
      [167, '[h]:mm:ss'],
    ],
    [164, 165, 166, 167],
  )

  for (let index = 0; index < 4; index++) {
    assert.equal(isDateFormat(parsed, index), true, `cell format ${index} should be a date`)
  }
})

test('does not mistake letters inside quoted text for date tokens', () => {
  const parsed = withFormats(
    [
      [164, '&quot;day&quot;0.0'],
      [165, '0.0&quot;m&quot;'],
    ],
    [164, 165],
  )

  assert.equal(isDateFormat(parsed, 0), false)
  assert.equal(isDateFormat(parsed, 1), false)
})

test('does not mistake a colour section for a date token', () => {
  const parsed = withFormats([[164, '[Red]#,##0.00']], [164])

  assert.equal(isDateFormat(parsed, 0), false)
})

test('does not mistake an escaped character for a date token', () => {
  const parsed = withFormats([[164, '0.0\\d']], [164])

  assert.equal(isDateFormat(parsed, 0), false)
})

test('treats General as a number', () => {
  const parsed = withFormats([], [0])

  assert.equal(isDateFormat(parsed, 0), false)
})

test('treats a cell with no style as a number', () => {
  const parsed = withFormats([], [14])

  assert.equal(isDateFormat(parsed, undefined), false)
})

test('treats a style index outside the table as a number', () => {
  const parsed = withFormats([], [14])

  assert.equal(isDateFormat(parsed, 99), false)
})

test('reads a styles part with no formats at all', () => {
  const parsed = readStyles(styles(''))

  assert.deepEqual(parsed.cellFormats, [])
  assert.equal(isDateFormat(parsed, 0), false)
})

test('reads the styles of every fixtures file', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  let parts = 0
  let dateFormats = 0

  for (const file of files) {
    const container = readContainer(new Uint8Array(await readFile(`fixtures/real/${file}`)))
    const part = container.parts.get('xl/styles.xml')
    if (part === undefined) continue

    const parsed = readStyles(new TextDecoder().decode(part))
    assert.ok(parsed.cellFormats.length > 0, `${file} has no cell formats`)

    for (let index = 0; index < parsed.cellFormats.length; index++) {
      if (isDateFormat(parsed, index)) dateFormats++
    }
    parts++
  }

  assert.ok(parts > 40, `expected many styles parts, got ${parts}`)
  assert.ok(dateFormats > 0, 'expected the fixtures to contain date formats')
})

test('ignores a number format with no code', () => {
  const parsed = readStyles(styles('<numFmts><numFmt numFmtId="164"/></numFmts>'))

  assert.equal(parsed.numberFormats.size, 0)
})

test('ignores a number format with an unreadable id', () => {
  const parsed = readStyles(
    styles('<numFmts><numFmt numFmtId="custom" formatCode="yyyy"/></numFmts>'),
  )

  assert.equal(parsed.numberFormats.size, 0)
})

test('defaults a cell format with no number format id to General', () => {
  const parsed = readStyles(styles('<cellXfs count="1"><xf fontId="0"/></cellXfs>'))

  assert.deepEqual(parsed.cellFormats, [0])
  assert.equal(isDateFormat(parsed, 0), false)
})

test('ignores cell style formats, which share the xf element name', () => {
  const parsed = readStyles(
    styles(
      '<cellStyleXfs count="1"><xf numFmtId="14"/></cellStyleXfs>' +
        '<cellXfs count="1"><xf numFmtId="0"/></cellXfs>',
    ),
  )

  assert.deepEqual(parsed.cellFormats, [0])
})

test('reads an empty cell format table written as a self closing element', () => {
  const parsed = readStyles(
    styles('<cellXfs count="0"/><cellStyleXfs><xf numFmtId="14"/></cellStyleXfs>'),
  )

  assert.deepEqual(parsed.cellFormats, [])
})

test('survives a format code with an unterminated bracket', () => {
  const parsed = withFormats([[164, '[Red#,##0']], [164])

  assert.equal(isDateFormat(parsed, 0), false)
})

test('treats an unknown format id as a number', () => {
  const parsed = withFormats([], [999])

  assert.equal(numberFormatOf(parsed, 0), undefined)
  assert.equal(isDateFormat(parsed, 0), false)
})
