import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { test } from 'node:test'
import { decodeXmlPart, readContainer, writeContainer } from './container.js'
import { XlsxError } from './errors.js'

async function fixtureFile(name: string): Promise<Uint8Array> {
  return new Uint8Array(await readFile(`fixtures/real/${name}`))
}

test('keeps every part when a file is read and written back', async () => {
  const original = await fixtureFile('WithChart.xlsx')

  const container = readContainer(original)
  const rewritten = readContainer(writeContainer(container))

  assert.deepEqual([...rewritten.parts.keys()].sort(), [...container.parts.keys()].sort())
})

test('keeps part contents byte for byte', async () => {
  const original = await fixtureFile('WithChart.xlsx')

  const container = readContainer(original)
  const rewritten = readContainer(writeContainer(container))

  for (const [path, bytes] of container.parts) {
    assert.deepEqual(rewritten.parts.get(path), bytes, `${path} changed`)
  }
})

test('reports the part that could not be read', async () => {
  const notAZip = new Uint8Array([1, 2, 3, 4])

  assert.throws(() => readContainer(notAZip), /not a zip/i)
})

test('writes the same bytes every time for the same parts', () => {
  const parts = new Map([['a.xml', new TextEncoder().encode('<a/>')]])

  const first = writeContainer({ parts })
  const header = first.slice(10, 14)

  // Local file header: bytes 10-11 hold the DOS time, 12-13 the DOS date.
  // A fixed stamp keeps output reproducible; 0x0021 is 1980-01-01.
  assert.deepEqual([...header], [0, 0, 33, 0])
})

test('write adds a part the package did not have, keeping the rest', () => {
  const source = readContainer(
    writeContainer({ parts: new Map([['a.xml', new TextEncoder().encode('<a/>')]]) }),
  )

  const out = source.write(new Map([['new.xml', new TextEncoder().encode('<new/>')]]))

  const parts = readContainer(out).parts
  assert.equal(new TextDecoder().decode(parts.get('a.xml')), '<a/>')
  assert.equal(new TextDecoder().decode(parts.get('new.xml')), '<new/>')
})

test('names encryption when handed an OLE2 compound file', () => {
  // A password-protected .xlsx and a legacy .xls are both OLE2, whose magic is
  // these eight bytes. The plain "not a zip" message sent users the wrong way.
  const ole2 = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00])

  assert.throws(
    () => readContainer(ole2),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'not-a-zip' &&
      /password|encrypt/i.test(error.message),
  )
})

test('decodes a normal utf-8 part, stripping a utf-8 BOM', () => {
  assert.equal(decodeXmlPart(new TextEncoder().encode('<a/>'), 'a.xml'), '<a/>')
  const withBom = new Uint8Array([0xef, 0xbb, 0xbf, 0x3c, 0x61, 0x2f, 0x3e])
  assert.equal(decodeXmlPart(withBom, 'a.xml'), '<a/>')
})

test('refuses a utf-16 part rather than decoding it to silent garbage', () => {
  // With a BOM the first byte is already invalid utf-8; without one, the nulls
  // decode cleanly and the sheet just looks empty. Both must be caught.
  const withBom = new Uint8Array([0xff, 0xfe, 0x3c, 0x00, 0x61, 0x00])
  const bomless = new Uint8Array([0x3c, 0x00, 0x61, 0x00, 0x2f, 0x00])

  for (const bytes of [withBom, bomless]) {
    assert.throws(
      () => decodeXmlPart(bytes, 'xl/worksheets/sheet1.xml'),
      (error: unknown) =>
        error instanceof XlsxError &&
        error.code === 'unreadable-part' &&
        error.part === 'xl/worksheets/sheet1.xml' &&
        /utf-?16|utf-8/i.test(error.message),
    )
  }
})

test('the parts of a read container behave as a full read-only map', () => {
  const source = writeContainer({
    parts: new Map([
      ['a.xml', new TextEncoder().encode('<a/>')],
      ['b.xml', new TextEncoder().encode('<b>two</b>')],
    ]),
  })
  const { parts } = readContainer(source)

  assert.equal(parts.size, 2)
  assert.deepEqual([...parts.keys()], ['a.xml', 'b.xml'])
  assert.equal(parts.has('a.xml'), true)
  assert.ok([...parts.values()].every((bytes) => bytes instanceof Uint8Array))
  // entries, the default iterator, and forEach must agree, and all inflate once.
  assert.deepEqual(
    [...parts.entries()].map(([name]) => name),
    [...parts].map(([name]) => name),
  )
  let visited = 0
  parts.forEach((bytes, name, map) => {
    visited++
    assert.equal(map.has(name), true)
    assert.ok(bytes.length > 0)
  })
  assert.equal(visited, 2)
})
