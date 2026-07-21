import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readContainer } from './container.js'
import { readSharedStrings } from './shared-strings.js'

const sst = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</sst>`

test('reads strings in table order', () => {
  const strings = readSharedStrings(sst('<si><t>first</t></si><si><t>second</t></si>'))

  assert.deepEqual(strings, ['first', 'second'])
})

test('joins the runs of a rich text string', () => {
  const strings = readSharedStrings(
    sst('<si><r><rPr><b/></rPr><t>bold</t></r><r><t> and plain</t></r></si>'),
  )

  assert.deepEqual(strings, ['bold and plain'])
})

test('keeps significant whitespace', () => {
  const strings = readSharedStrings(sst('<si><t xml:space="preserve">  padded  </t></si>'))

  assert.deepEqual(strings, ['  padded  '])
})

test('ignores phonetic runs', () => {
  const strings = readSharedStrings(
    sst('<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>'),
  )

  assert.deepEqual(strings, ['東京'])
})

test('decodes entities', () => {
  const strings = readSharedStrings(sst('<si><t>a &amp; b</t></si>'))

  assert.deepEqual(strings, ['a & b'])
})

test('reads an empty string entry', () => {
  const strings = readSharedStrings(sst('<si><t></t></si><si><t>x</t></si>'))

  assert.deepEqual(strings, ['', 'x'])
})

test('reads a table with no entries', () => {
  assert.deepEqual(readSharedStrings(sst('')), [])
})

test('reads the shared strings of every fixtures file that has them', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  let tables = 0
  let entries = 0

  for (const file of files) {
    const container = readContainer(new Uint8Array(await readFile(`fixtures/real/${file}`)))
    const part = container.parts.get('xl/sharedStrings.xml')
    if (part === undefined) continue

    const strings = readSharedStrings(new TextDecoder().decode(part))
    const declared = /uniqueCount="(\d+)"/.exec(new TextDecoder().decode(part))?.[1]

    if (declared !== undefined) {
      assert.equal(
        strings.length,
        Number(declared),
        `${file}: read ${strings.length} strings but the table declares ${declared}`,
      )
    }
    tables++
    entries += strings.length
  }

  assert.ok(tables > 10, `expected many shared string tables, got ${tables}`)
  assert.ok(entries > 100, `expected many strings, got ${entries}`)
})
