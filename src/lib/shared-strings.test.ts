import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { assertWellFormed } from '../testing/invariants.js'
import { readContainer } from './container.js'
import { appendSharedStrings, readSharedStrings } from './shared-strings.js'

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

const sstOf = (result: { xml: string }) => result.xml

test('returns the index of a string already in the table', () => {
  const result = appendSharedStrings(sst('<si><t>alpha</t></si><si><t>beta</t></si>'), ['beta'])

  assert.equal(result.indexes.get('beta'), 1)
  assert.equal(sstOf(result).includes('<si><t>beta</t></si><si>'), false)
})

test('appends a string that is not in the table yet', () => {
  const result = appendSharedStrings(sst('<si><t>alpha</t></si>'), ['gamma'])

  assert.equal(result.indexes.get('gamma'), 1)
  assert.match(sstOf(result), /<si><t>alpha<\/t><\/si><si><t>gamma<\/t><\/si><\/sst>/)
})

test('appends several strings in the order given', () => {
  const result = appendSharedStrings(sst(''), ['one', 'two'])

  assert.equal(result.indexes.get('one'), 0)
  assert.equal(result.indexes.get('two'), 1)
  assert.match(sstOf(result), /<si><t>one<\/t><\/si><si><t>two<\/t><\/si>/)
})

test('appends a repeated string only once', () => {
  const result = appendSharedStrings(sst(''), ['same', 'same'])

  assert.equal(result.indexes.get('same'), 0)
  assert.equal(sstOf(result).match(/<si>/g)?.length, 1)
})

test('updates the unique count', () => {
  const source = `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="1"><si><t>alpha</t></si></sst>`

  const result = appendSharedStrings(source, ['beta'])

  assert.match(sstOf(result), /uniqueCount="2"/)
})

test('raises the total count by the number of strings added', () => {
  const source = `<sst count="3" uniqueCount="1"><si><t>alpha</t></si></sst>`

  const result = appendSharedStrings(source, ['beta', 'gamma'])

  assert.match(sstOf(result), /count="5"/)
})

test('escapes a string that would break the markup', () => {
  const result = appendSharedStrings(sst(''), ['a & b < c'])

  assert.match(sstOf(result), /<si><t>a &amp; b &lt; c<\/t><\/si>/)
})

test('preserves whitespace that would otherwise be trimmed', () => {
  const result = appendSharedStrings(sst(''), ['  padded  '])

  assert.match(sstOf(result), /<t xml:space="preserve"> {2}padded {2}<\/t>/)
})

test('leaves the entries already in the table untouched', () => {
  const source = sst('<si><r><t>rich</t></r><r><t> text</t></r></si>')

  const result = appendSharedStrings(source, ['new'])

  assert.match(sstOf(result), /<si><r><t>rich<\/t><\/r><r><t> text<\/t><\/r><\/si>/)
})

test('writes into a table with no entries at all', () => {
  const result = appendSharedStrings('<sst/>', ['first'])

  assert.equal(result.indexes.get('first'), 0)
  assert.match(sstOf(result), /<sst[^>]*><si><t>first<\/t><\/si><\/sst>/)
})

test('changes nothing when every string is already there', () => {
  const source = sst('<si><t>alpha</t></si>')

  const result = appendSharedStrings(source, ['alpha'])

  assert.equal(sstOf(result), source)
})

test('maps a duplicated entry to the first place it appears', () => {
  const result = appendSharedStrings(
    sst('<si><t>same</t></si><si><t>other</t></si><si><t>same</t></si>'),
    ['same'],
  )

  assert.equal(result.indexes.get('same'), 0)
})

test('keeps the declaration when the table was self closing', () => {
  const result = appendSharedStrings('<?xml version="1.0"?><sst/>', ['x'])

  assert.match(result.xml, /^<\?xml version="1\.0"\?><sst><si><t>x<\/t><\/si><\/sst>$/)
})

test('rejects text that xml cannot represent', () => {
  assert.throws(
    () => appendSharedStrings(sst(''), [`a${String.fromCharCode(7)}b`]),
    /cannot be written to xml/i,
  )
})

test('escapes a carriage return in a shared string', () => {
  const result = appendSharedStrings('<sst count="0" uniqueCount="0"></sst>', ['a\r\nb'])

  assert.equal(result.xml.includes('\r'), false)
  assert.match(result.xml, /a&#13;\nb/)
})

test('writes new entries with the prefix the table uses', () => {
  const source =
    '<x:sst xmlns:x="http://x" count="1" uniqueCount="1"><x:si><x:t>a</x:t></x:si></x:sst>'

  const result = appendSharedStrings(source, ['b'])

  assertWellFormed(result.xml, 'prefixed sst')
  assert.match(result.xml, /<x:si><x:t>b<\/x:t><\/x:si>/)
})

test('closes a self closing prefixed table with the matching tag', () => {
  const result = appendSharedStrings('<x:sst xmlns:x="http://x"/>', ['b'])

  assertWellFormed(result.xml, 'self closing prefixed sst')
  assert.equal(result.xml.includes('</sst>'), false)
})

test('counts references even when uniqueCount comes first', () => {
  const source = '<sst uniqueCount="1" count="1"><si><t>a</t></si></sst>'

  const result = appendSharedStrings(source, ['b'])

  assert.match(result.xml, /count="2"/)
  assert.match(result.xml, /uniqueCount="2"/)
})

test('does not reuse an entry whose text is assembled from formatted runs', () => {
  // Reusing it would give the newly written cell the bold run's formatting,
  // which nothing the caller did asked for.
  const xml = sst('<si><r><rPr><b/></rPr><t>To</t></r><r><t>tal</t></r></si>')

  const appended = appendSharedStrings(xml, ['Total'])

  assert.equal(appended.indexes.get('Total'), 1)
  assert.match(appended.xml, /<si><t>Total<\/t><\/si>/)
})

test('does not reuse an entry carrying a phonetic guide', () => {
  const xml = sst('<si><t>Total</t><rPh sb="0" eb="2"><t>guide</t></rPh></si>')

  assert.equal(appendSharedStrings(xml, ['Total']).indexes.get('Total'), 1)
})
