import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { XlsxError } from './errors.js'
import { readSheet } from './sheet.js'
import { readSharedStrings } from './shared-strings.js'
import { readWorkbookPart } from './workbook.js'

const sheet = (body: string, dimension = 'A1:Z100') =>
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${dimension}"/><sheetData>${body}</sheetData></worksheet>`

const encode = (xml: string) => new TextEncoder().encode(xml)

const cells = (body: string, strings: readonly string[] = []) => [
  ...readSheet(encode(sheet(body)), strings),
]

test('reads a sheet from utf-8 bytes, multibyte content and all', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="inlineStr"><is><t>café €</t></is></c></row>')

  assert.deepEqual(cell?.value, { kind: 'text', value: 'café €' })
})

test('reads a number', () => {
  const [cell] = cells('<row r="1"><c r="A1"><v>42.5</v></c></row>')

  assert.equal(cell?.reference, 'A1')
  assert.deepEqual(cell?.value, { kind: 'number', value: 42.5 })
})

test('reads a shared string by index', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="s"><v>1</v></c></row>', ['zero', 'one'])

  assert.deepEqual(cell?.value, { kind: 'text', value: 'one' })
})

test('reads an inline string', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c></row>')

  assert.deepEqual(cell?.value, { kind: 'text', value: 'hello' })
})

test('reads a formula result typed as a string', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="str"><f>A2</f><v>done</v></c></row>')

  assert.deepEqual(cell?.value, { kind: 'text', value: 'done' })
  assert.equal(cell?.formula, 'A2')
})

test('reads booleans', () => {
  const [yes, no] = cells(
    '<row r="1"><c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c></row>',
  )

  assert.deepEqual(yes?.value, { kind: 'boolean', value: true })
  assert.deepEqual(no?.value, { kind: 'boolean', value: false })
})

test('reads error values', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="e"><v>#DIV/0!</v></c></row>')

  assert.deepEqual(cell?.value, { kind: 'error', value: '#DIV/0!' })
})

test('reads a formula and its cached result', () => {
  const [cell] = cells('<row r="1"><c r="A1"><f>SUM(B1:B3)</f><v>6</v></c></row>')

  assert.equal(cell?.formula, 'SUM(B1:B3)')
  assert.deepEqual(cell?.value, { kind: 'number', value: 6 })
})

test('keeps the style index so formatting can be resolved later', () => {
  const [cell] = cells('<row r="1"><c r="A1" s="3"><v>1</v></c></row>')

  assert.equal(cell?.styleIndex, 3)
})

test('reports a cell with no value as empty', () => {
  const [cell] = cells('<row r="1"><c r="A1" s="2"/></row>')

  assert.deepEqual(cell?.value, { kind: 'empty' })
})

test('places cells that carry no reference by position', () => {
  const found = cells('<row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row>')

  assert.deepEqual(
    found.map((cell) => cell.reference),
    ['A1', 'B1', 'A2'],
  )
})

test('reads rows that appear out of order', () => {
  const found = cells(
    '<row r="3"><c r="A3"><v>300</v></c></row><row r="1"><c r="A1"><v>100</v></c></row>',
  )

  assert.deepEqual(
    found.map((cell) => cell.reference),
    ['A3', 'A1'],
  )
})

test('reads cells beyond the range the dimension claims', () => {
  const found = [
    ...readSheet(encode(sheet('<row r="1"><c r="C3"><v>9</v></c></row>', 'A1:A1')), []),
  ]

  assert.deepEqual(
    found.map((cell) => cell.reference),
    ['C3'],
  )
})

test('reads columns past Z', () => {
  const found = cells('<row r="1"><c r="Z1"><v>26</v></c><c r="AA1"><v>27</v></c></row>')

  assert.deepEqual(
    found.map((cell) => cell.address.column),
    [26, 27],
  )
})

test('rejects a shared string index the table does not hold', () => {
  // The file is at fault, not the caller: a t="s" cell must point at a string the
  // shared table actually holds. A missing one is corrupt content, not empty text.
  const rejects = (raw: string) =>
    assert.throws(
      () => cells(`<row r="1"><c r="A1" t="s"><v>${raw}</v></c></row>`, ['only']),
      (error: unknown) =>
        error instanceof XlsxError && error.code === 'invalid-content' && error.reference === 'A1',
    )
  rejects('9') // past the end of the table
  rejects('-1')
  rejects('abc')
})

test('rejects a number it cannot read', () => {
  // The file is at fault, not the caller: nobody was trying to write anything.
  assert.throws(
    () => cells('<row r="1"><c r="A1"><v>banana</v></c></row>'),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'invalid-content' && error.reference === 'A1',
  )
})

test('reads every sheet in the fixtures', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  const failures: string[] = []
  let total = 0

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const workbook = readWorkbookPart(bytes)

    const stringsPart = workbook.container.parts.get('xl/sharedStrings.xml')
    const strings =
      stringsPart === undefined ? [] : readSharedStrings(new TextDecoder().decode(stringsPart))

    for (const sheetEntry of workbook.sheets) {
      const part = workbook.container.parts.get(sheetEntry.path)
      if (part === undefined) continue
      try {
        for (const _cell of readSheet(part, strings)) total++
      } catch (error) {
        failures.push(
          `${file} ${sheetEntry.name}: ${error instanceof Error ? error.message : String(error)}`,
        )
      }
    }
  }

  assert.deepEqual(failures, [])
  assert.ok(total > 5000, `expected many cells across the fixtures, got ${total}`)
})

test('reads a shared formula dependent, which carries no expression', () => {
  const [master, dependent] = cells(
    '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1*2</f><v>2</v></c></row>' +
      '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>',
  )

  assert.equal(master?.formula, 'B1*2')
  assert.equal(dependent?.formula, '')
  assert.deepEqual(dependent?.value, { kind: 'number', value: 4 })
})

test('reads a cell whose value element is self closing', () => {
  const [cell] = cells('<row r="1"><c r="A1"><v/></c></row>')

  assert.deepEqual(cell?.value, { kind: 'empty' })
})

test('reads an inline string with a self closing text element', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="inlineStr"><is><t/></is></c></row>')

  assert.deepEqual(cell?.value, { kind: 'text', value: '' })
})

test('numbers rows that carry no reference after one that does', () => {
  const found = cells('<row r="5"><c><v>1</v></c></row><row><c><v>2</v></c></row>')

  assert.deepEqual(
    found.map((cell) => cell.reference),
    ['A5', 'A6'],
  )
})

test('reads an ISO date cell rather than failing on it', () => {
  const [cell] = cells('<row r="1"><c r="A1" t="d"><v>2024-01-01T00:00:00</v></c></row>')

  assert.deepEqual(cell?.value, { kind: 'date', value: '2024-01-01T00:00:00' })
})

test('does not take phonetic guides as inline string content', () => {
  const [cell] = cells(
    '<row r="1"><c r="A1" t="inlineStr"><is><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh></is></c></row>',
  )

  assert.deepEqual(cell?.value, { kind: 'text', value: '東京' })
})

test('reads a sheet whose elements carry a namespace prefix', () => {
  const xml =
    '<x:worksheet><x:sheetData><x:row r="1">' +
    '<x:c r="A1"><x:v>1</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>hi</x:t></x:is></x:c>' +
    '</x:row></x:sheetData></x:worksheet>'

  const found = [...readSheet(encode(xml), [])]

  assert.deepEqual(
    found.map((cell) => cell.reference),
    ['A1', 'B1'],
  )
  assert.deepEqual(found[1]?.value, { kind: 'text', value: 'hi' })
})
