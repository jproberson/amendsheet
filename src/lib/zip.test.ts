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

test('reports a part too large to hold in memory rather than crashing', () => {
  const huge: ZipEntry = {
    name: 'xl/worksheets/sheet1.xml',
    method: 8,
    crc: 0,
    compressed: enc('x'),
    uncompressedSize: 2 ** 48,
  }
  assert.throws(
    () => inflate(huge),
    (error: unknown) =>
      error instanceof XlsxError &&
      error.code === 'part-too-large' &&
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

test('rejects a central-directory offset that points past the archive', () => {
  const bytes = corrupted((view, _b, end) => view.setUint32(end + 16, 0x7fffff00, true))
  assert.throws(() => readZip(bytes), notAZipMatching(/central/i))
})

test('rejects a local-header offset that points past the archive', () => {
  const bytes = corrupted((view, _b, end) => {
    const central = view.getUint32(end + 16, true)
    view.setUint32(central + 42, 0x7fffff00, true)
  })
  assert.throws(() => readZip(bytes), notAZipMatching(/local header/i))
})

function unreachable(): never {
  throw new Error('entry was missing')
}

/** An archive with `count` tiny stored entries, enough to force a ZIP64 count. */
function manyEntryArchive(count: number): Uint8Array {
  const entries: ZipEntry[] = []
  for (let index = 0; index < count; index++) {
    entries.push({
      name: `p${index}.xml`,
      method: 0,
      crc: 0,
      compressed: enc('x'),
      uncompressedSize: 1,
    })
  }
  return writeZip(entries)
}

test('round-trips an archive with more than 65535 entries (ZIP64 count)', () => {
  const count = 70000
  const read = readZip(manyEntryArchive(count))

  assert.equal(read.length, count)
  assert.equal(read[0]?.name, 'p0.xml')
  assert.equal(read[count - 1]?.name, `p${count - 1}.xml`)
  assert.equal(dec(inflate(read[12345] ?? unreachable())), 'x')
})

test('round-trips an archive of exactly 65535 entries', () => {
  const count = 65535
  const read = readZip(manyEntryArchive(count))

  assert.equal(read.length, count)
  assert.equal(read[count - 1]?.name, `p${count - 1}.xml`)
})

test('refuses a ZIP64 count whose end-of-central-directory record is corrupt', () => {
  const bytes = manyEntryArchive(65536)
  // The ZIP64 record sits before the 20-byte locator and the 22-byte end record.
  const record = bytes.length - 22 - 20 - 56
  new DataView(bytes.buffer).setUint32(record, 0, true) // wipe its signature
  assert.throws(() => readZip(bytes), notAZipMatching(/ZIP64 end-of-central-directory/))
})

/** A minimal archive whose one entry declares 32-bit overflow and carries a
 * ZIP64 extra field, exactly as a >4GB file's directory would — with small real
 * values, so the read path is tested without a 4GB allocation. */
function zip64EntryArchive(options: {
  sizeOverflow: boolean
  offsetOverflow: boolean
}): Uint8Array {
  const name = enc('big.bin')
  const data = enc('hello')
  const U32 = 0xffffffff

  const localExtra = options.sizeOverflow ? 20 : 0
  const local = new Uint8Array(30 + name.length + localExtra + data.length)
  const lv = new DataView(local.buffer)
  lv.setUint32(0, 0x04034b50, true)
  lv.setUint16(8, 0, true) // stored
  lv.setUint32(18, options.sizeOverflow ? U32 : data.length, true)
  lv.setUint32(22, options.sizeOverflow ? U32 : data.length, true)
  lv.setUint16(26, name.length, true)
  lv.setUint16(28, localExtra, true)
  local.set(name, 30)
  if (options.sizeOverflow) {
    lv.setUint16(30 + name.length, 0x0001, true)
    lv.setUint16(30 + name.length + 2, 16, true)
    lv.setBigUint64(30 + name.length + 4, BigInt(data.length), true)
    lv.setBigUint64(30 + name.length + 12, BigInt(data.length), true)
  }
  local.set(data, 30 + name.length + localExtra)

  const extraFields = (options.sizeOverflow ? 16 : 0) + (options.offsetOverflow ? 8 : 0)
  const centralExtra = extraFields > 0 ? 4 + extraFields : 0
  const central = new Uint8Array(46 + name.length + centralExtra)
  const cv = new DataView(central.buffer)
  cv.setUint32(0, 0x02014b50, true)
  cv.setUint16(10, 0, true)
  cv.setUint32(20, options.sizeOverflow ? U32 : data.length, true)
  cv.setUint32(24, options.sizeOverflow ? U32 : data.length, true)
  cv.setUint16(28, name.length, true)
  cv.setUint16(30, centralExtra, true)
  cv.setUint32(42, options.offsetOverflow ? U32 : 0, true)
  central.set(name, 46)
  if (centralExtra > 0) {
    cv.setUint16(46 + name.length, 0x0001, true)
    cv.setUint16(46 + name.length + 2, extraFields, true)
    let at = 46 + name.length + 4
    if (options.sizeOverflow) {
      cv.setBigUint64(at, BigInt(data.length), true)
      cv.setBigUint64(at + 8, BigInt(data.length), true)
      at += 16
    }
    if (options.offsetOverflow) cv.setBigUint64(at, 0n, true) // real local offset is 0
  }

  const end = new Uint8Array(22)
  const ev = new DataView(end.buffer)
  ev.setUint32(0, 0x06054b50, true)
  ev.setUint16(8, 1, true)
  ev.setUint16(10, 1, true)
  ev.setUint32(12, central.length, true)
  ev.setUint32(16, local.length, true)

  const out = new Uint8Array(local.length + central.length + end.length)
  out.set(local, 0)
  out.set(central, local.length)
  out.set(end, local.length + central.length)
  return out
}

test('reads an entry whose sizes overflowed into a ZIP64 extra field', () => {
  const [entry] = readZip(zip64EntryArchive({ sizeOverflow: true, offsetOverflow: false }))
  assert.equal(entry?.name, 'big.bin')
  assert.equal(entry?.uncompressedSize, 5)
  assert.equal(dec(inflate(entry ?? unreachable())), 'hello')
})

test('reads an entry whose local offset overflowed into a ZIP64 extra field', () => {
  const [entry] = readZip(zip64EntryArchive({ sizeOverflow: false, offsetOverflow: true }))
  assert.equal(dec(inflate(entry ?? unreachable())), 'hello')
})

test('reads an entry whose size and offset both overflowed', () => {
  const [entry] = readZip(zip64EntryArchive({ sizeOverflow: true, offsetOverflow: true }))
  assert.equal(dec(inflate(entry ?? unreachable())), 'hello')
})

test('refuses an entry that maxes a field but carries no ZIP64 extra', () => {
  const bytes = zip64EntryArchive({ sizeOverflow: true, offsetOverflow: false })
  // Zero the central header's extra-length field so the ZIP64 field vanishes.
  const centralStart = bytes.length - 22 - (46 + enc('big.bin').length + 20)
  new DataView(bytes.buffer).setUint16(centralStart + 30, 0, true)
  assert.throws(
    () => readZip(bytes),
    (error: unknown) =>
      error instanceof XlsxError && error.code === 'not-a-zip' && /ZIP64 extra/.test(error.message),
  )
})

test('refuses a ZIP64 directory with no locator before the end record', () => {
  const bytes = writeZip([
    { name: 'a', method: 0, crc: 0, compressed: enc('x'), uncompressedSize: 1 },
  ])
  // Force the plain end record to claim ZIP64 without any ZIP64 structures.
  new DataView(bytes.buffer).setUint16(endOffset(bytes) + 10, 0xffff, true)
  assert.throws(() => readZip(bytes), notAZipMatching(/no ZIP64 locator/))
})

test('refuses a value beyond what a JS number can hold exactly', () => {
  const bytes = zip64EntryArchive({ sizeOverflow: true, offsetOverflow: false })
  const centralStart = bytes.length - 22 - (46 + enc('big.bin').length + 20)
  // The uncompressed-size slot in the central ZIP64 extra: 9 quadrillion + 1.
  const extraBody = centralStart + 46 + enc('big.bin').length + 4
  new DataView(bytes.buffer).setBigUint64(extraBody, BigInt(Number.MAX_SAFE_INTEGER) + 2n, true)
  assert.throws(() => readZip(bytes), notAZipMatching(/9PB/))
})

test('writes and reads back an entry declaring a size past 4GB', () => {
  // A pretend-huge entry: real bytes are tiny, but uncompressedSize crosses the
  // 32-bit line, which forces the per-entry ZIP64 extra field on the write path.
  const huge: ZipEntry = {
    name: 'big',
    method: 0,
    crc: 0,
    compressed: enc('x'),
    uncompressedSize: 0x1_0000_0000,
  }
  const [entry] = readZip(writeZip([huge]))
  assert.equal(entry?.uncompressedSize, 0x1_0000_0000)
  assert.equal(dec(inflate(entry ?? unreachable())), 'x')
})
