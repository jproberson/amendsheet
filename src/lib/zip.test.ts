import assert from 'node:assert/strict'
import { test } from 'node:test'
import { XlsxError } from './errors.js'
import { type ZipEntry, deflate, inflate, readZip, writeZip } from './zip.js'

const enc = (text: string) => new TextEncoder().encode(text)
const dec = (bytes: Uint8Array) => new TextDecoder().decode(bytes)

const valid = () => writeZip([deflate('a.xml', enc('<a/>')), deflate('b.xml', enc('<b>x</b>'))])

const END_SIG = 0x06054b50
function endOffset(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let offset = bytes.length - 22; offset >= 0; offset--) {
    if (view.getUint32(offset, true) === END_SIG) return offset
  }
  throw new Error('no end record')
}

/** A valid archive with one field overwritten, to reach a rejection path. */
function corrupted(mutate: (view: DataView, bytes: Uint8Array, end: number) => void): Uint8Array {
  const bytes = valid().slice()
  const view = new DataView(bytes.buffer)
  mutate(view, bytes, endOffset(bytes))
  return bytes
}

const isNotAZip = (error: unknown) => error instanceof XlsxError && error.code === 'not-a-zip'

const notAZipMatching = (pattern: RegExp) => (error: unknown) =>
  error instanceof XlsxError && error.code === 'not-a-zip' && pattern.test(error.message)

test('round-trips entries through write and read, inflating on demand', () => {
  const entries = readZip(valid())

  assert.deepEqual(
    entries.map((entry) => entry.name),
    ['a.xml', 'b.xml'],
  )
  assert.equal(dec(inflate(entries[0] ?? unreachable())), '<a/>')
  assert.equal(dec(inflate(entries[1] ?? unreachable())), '<b>x</b>')
})

test('copies a stored entry through rather than inflating it', () => {
  const stored: ZipEntry = {
    name: 's',
    method: 0,
    crc: 0,
    compressed: enc('raw'),
    uncompressedSize: 3,
  }
  assert.equal(dec(inflate(stored)), 'raw')
})

test('refuses an entry compressed with a method it does not know', () => {
  const entry: ZipEntry = {
    name: 's',
    method: 99,
    crc: 0,
    compressed: enc('x'),
    uncompressedSize: 1,
  }
  assert.throws(
    () => inflate(entry),
    (error: unknown) => error instanceof XlsxError && error.code === 'unreadable-part',
  )
})

test('reports a part that cannot be inflated', () => {
  const broken: ZipEntry = {
    name: 'xl/worksheets/sheet1.xml',
    method: 8,
    crc: 0,
    compressed: enc('not a deflate stream'),
    uncompressedSize: 100,
  }
  assert.throws(
    () => inflate(broken),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'unreadable-part' &&
      error.part === 'xl/worksheets/sheet1.xml',
  )
})

test('rejects bytes with no end-of-central-directory record', () => {
  assert.throws(() => readZip(enc('nowhere near a zip archive')), isNotAZip)
})

test('rejects a ZIP64 archive rather than misreading it', () => {
  const manyEntries = corrupted((view, _bytes, end) => view.setUint16(end + 10, 0xffff, true))
  const hugeOffset = corrupted((view, _bytes, end) => view.setUint32(end + 16, 0xffffffff, true))

  assert.throws(() => readZip(manyEntries), notAZipMatching(/ZIP64/))
  assert.throws(() => readZip(hugeOffset), isNotAZip)
})

test('rejects a ZIP64 marker inside a central directory entry', () => {
  const bytes = corrupted((view, _b, end) => {
    const central = view.getUint32(end + 16, true)
    view.setUint32(central + 20, 0xffffffff, true) // compressed size
  })
  assert.throws(() => readZip(bytes), notAZipMatching(/ZIP64/))
})

test('rejects a malformed central directory header', () => {
  const bytes = corrupted((view, _b, end) =>
    view.setUint32(view.getUint32(end + 16, true), 0, true),
  )
  assert.throws(() => readZip(bytes), notAZipMatching(/central/i))
})

test('rejects a malformed local header', () => {
  const bytes = corrupted((view, _b, end) => {
    const central = view.getUint32(end + 16, true)
    const localOffset = view.getUint32(central + 42, true)
    view.setUint32(localOffset, 0, true)
  })
  assert.throws(() => readZip(bytes), notAZipMatching(/local header/i))
})

function unreachable(): never {
  throw new Error('entry was missing')
}
