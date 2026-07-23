import { deflateSync, inflateSync } from 'fflate'
import { XlsxError } from './errors.js'

/**
 * Just enough of the ZIP format to own the container: read entries without
 * inflating them, and write an archive that mixes freshly-compressed parts with
 * ones copied through still-compressed. fflate stays, but only for the DEFLATE
 * of the parts we actually change.
 */
export interface ZipEntry {
  readonly name: string
  /** 0 is stored, 8 is DEFLATE. Nothing else is written into an .xlsx. */
  readonly method: number
  readonly crc: number
  readonly compressed: Uint8Array
  readonly uncompressedSize: number
}

const LOCAL_SIG = 0x04034b50
const CENTRAL_SIG = 0x02014b50
const END_SIG = 0x06054b50
const ZIP64_EOCD_SIG = 0x06064b50
const ZIP64_LOCATOR_SIG = 0x07064b50
/** Header id of the ZIP64 extended-information field inside an entry's extra. */
const ZIP64_EXTRA_ID = 0x0001
/** A 32-bit field holding this value means the real value is in a ZIP64 record. */
const U32_MAX = 0xffffffff
const U16_MAX = 0xffff

const notAZip = (message: string) => new XlsxError('not-a-zip', message, {})

/** Reads an 8-byte field, refusing a value too large for a JS number to hold exactly. */
function readU64(view: DataView, offset: number): number {
  const value = view.getBigUint64(offset, true)
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw notAZip('Archive declares a size beyond 9PB, which cannot be represented')
  }
  return Number(value)
}

function endOfCentralDirectory(view: DataView, length: number): number {
  // The record is 22 bytes plus an optional trailing comment, so it is found by
  // scanning back from the end for its signature.
  for (let offset = length - 22; offset >= 0; offset--) {
    if (view.getUint32(offset, true) === END_SIG) return offset
  }
  throw notAZip('File has no end-of-central-directory record, so it is not a zip')
}

/**
 * The count and central-directory offset, from the ZIP64 record when the plain
 * end-of-central-directory maxes out either field, otherwise from the plain one.
 */
function directoryStart(view: DataView, end: number): { count: number; offset: number } {
  const count = view.getUint16(end + 10, true)
  const offset = view.getUint32(end + 16, true)
  if (count !== U16_MAX && offset !== U32_MAX) return { count, offset }

  // The ZIP64 locator sits immediately before the plain record and points at
  // the ZIP64 end-of-central-directory record, which holds the real values.
  const locator = end - 20
  if (locator < 0 || view.getUint32(locator, true) !== ZIP64_LOCATOR_SIG) {
    throw notAZip('Archive maxes out a directory field but has no ZIP64 locator')
  }
  const record = readU64(view, locator + 8)
  if (
    record < 0 ||
    record + 56 > view.byteLength ||
    view.getUint32(record, true) !== ZIP64_EOCD_SIG
  ) {
    throw notAZip('ZIP64 end-of-central-directory record is missing or malformed')
  }
  return { count: readU64(view, record + 32), offset: readU64(view, record + 48) }
}

/**
 * Reads the values a ZIP64 extra field holds for the fields that overflowed,
 * which appear in the order uncompressed size, compressed size, local offset —
 * only the ones whose 32-bit field was maxed out.
 */
function readZip64Extra(view: DataView, start: number, length: number, name: string): () => number {
  let position = start
  const end = start + length
  while (position + 4 <= end) {
    const id = view.getUint16(position, true)
    const size = view.getUint16(position + 2, true)
    if (id === ZIP64_EXTRA_ID) {
      let cursor = position + 4
      return () => {
        const value = readU64(view, cursor)
        cursor += 8
        return value
      }
    }
    position += 4 + size
  }
  throw notAZip(`Entry ${name} needs a ZIP64 extra field but has none`)
}

/** Reads every entry's metadata and raw bytes, inflating nothing. */
export function readZip(bytes: Uint8Array): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = endOfCentralDirectory(view, bytes.length)
  const { count, offset: centralOffset } = directoryStart(view, end)

  const entries: ZipEntry[] = []
  let offset = centralOffset
  for (let index = 0; index < count; index++) {
    // The offset comes from the file, so it can point past the buffer; reading
    // there throws a bare RangeError unless the fixed 46-byte header is in range.
    if (
      offset < 0 ||
      offset + 46 > view.byteLength ||
      view.getUint32(offset, true) !== CENTRAL_SIG
    ) {
      throw notAZip('Central directory is malformed')
    }
    const method = view.getUint16(offset + 10, true)
    const crc = view.getUint32(offset + 16, true)
    let compressedSize = view.getUint32(offset + 20, true)
    let uncompressedSize = view.getUint32(offset + 24, true)
    const nameLength = view.getUint16(offset + 28, true)
    const extraLength = view.getUint16(offset + 30, true)
    const commentLength = view.getUint16(offset + 32, true)
    let localOffset = view.getUint32(offset + 42, true)

    const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength))

    if (uncompressedSize === U32_MAX || compressedSize === U32_MAX || localOffset === U32_MAX) {
      const next = readZip64Extra(view, offset + 46 + nameLength, extraLength, name)
      if (uncompressedSize === U32_MAX) uncompressedSize = next()
      if (compressedSize === U32_MAX) compressedSize = next()
      if (localOffset === U32_MAX) localOffset = next()
    }

    if (
      localOffset < 0 ||
      localOffset + 30 > view.byteLength ||
      view.getUint32(localOffset, true) !== LOCAL_SIG
    ) {
      throw notAZip(`Local header for ${name} is malformed`)
    }
    // The local header repeats the name and extra fields; the data follows them,
    // and their lengths there may differ from the central directory's.
    const localNameLength = view.getUint16(localOffset + 26, true)
    const localExtraLength = view.getUint16(localOffset + 28, true)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength

    entries.push({
      name,
      method,
      crc,
      uncompressedSize,
      compressed: bytes.subarray(dataStart, dataStart + compressedSize),
    })
    offset += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

const STORED = 0
const DEFLATED = 8

/** Returns an entry's content, inflating a DEFLATE entry and copying a stored one. */
export function inflate(entry: ZipEntry): Uint8Array {
  if (entry.method === STORED) return entry.compressed
  if (entry.method === DEFLATED) {
    let out: Uint8Array
    try {
      out = new Uint8Array(entry.uncompressedSize)
    } catch (cause) {
      throw new XlsxError(
        'part-too-large',
        `Part ${entry.name} decompresses to ${entry.uncompressedSize} bytes, too large to hold in memory`,
        { part: entry.name, cause },
      )
    }
    try {
      return inflateSync(entry.compressed, { out })
    } catch (cause) {
      throw new XlsxError('unreadable-part', `Part ${entry.name} could not be inflated`, {
        part: entry.name,
        cause,
      })
    }
  }
  throw new XlsxError('unreadable-part', `Part ${entry.name} uses compression ${entry.method}`, {
    part: entry.name,
  })
}

/** Branchless so there is no table lookup to guard, and the polynomial is the
 * standard CRC-32 one ZIP uses. */
function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of bytes) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1))
  }
  return (crc ^ 0xffffffff) >>> 0
}

/** Compresses a part into an entry that can be written or passed through. */
export function deflate(name: string, bytes: Uint8Array): ZipEntry {
  return {
    name,
    method: DEFLATED,
    crc: crc32(bytes),
    compressed: deflateSync(bytes, { level: 6 }),
    uncompressedSize: bytes.length,
  }
}

// DOS time 0 and date 0x0021 are 1980-01-01: a fixed stamp keeps output stable.
const DOS_TIME = 0
const DOS_DATE = 0x0021
const VERSION = 20
const VERSION_ZIP64 = 45

const setU64 = (view: DataView, offset: number, value: number) =>
  view.setBigUint64(offset, BigInt(value), true)

interface Placed {
  readonly entry: ZipEntry
  readonly name: Uint8Array
  /** Its compressed or uncompressed size needs 64 bits, so the entry is ZIP64. */
  readonly sizeZip64: boolean
  readonly localOffset: number
  /** Its local-header offset needs 64 bits (an archive past 4GB). */
  readonly offsetZip64: boolean
  /** Bytes of the central-header ZIP64 extra field, 0 when the entry is plain. */
  readonly centralExtra: number
}

/**
 * Writes the archive from entries that already carry their compressed bytes.
 * A part, an offset or the entry count that will not fit in 32 or 16 bits emits
 * the ZIP64 records the format defines for it; a package within those limits —
 * every ordinary workbook — comes out byte for byte as it did before ZIP64.
 */
export function writeZip(entries: readonly ZipEntry[]): Uint8Array {
  const encoder = new TextEncoder()

  let localTotal = 0
  const placed: Placed[] = []
  for (const entry of entries) {
    const name = encoder.encode(entry.name)
    const sizeZip64 = entry.compressed.length >= U32_MAX || entry.uncompressedSize >= U32_MAX
    const localOffset = localTotal
    localTotal += 30 + name.length + (sizeZip64 ? 20 : 0) + entry.compressed.length
    placed.push({ entry, name, sizeZip64, localOffset, offsetZip64: false, centralExtra: 0 })
  }

  const centralStart = localTotal
  let centralTotal = 0
  const resolved = placed.map((item): Placed => {
    const offsetZip64 = item.localOffset >= U32_MAX
    const centralExtra =
      item.sizeZip64 || offsetZip64 ? 4 + (item.sizeZip64 ? 16 : 0) + (offsetZip64 ? 8 : 0) : 0
    centralTotal += 46 + item.name.length + centralExtra
    return { ...item, offsetZip64, centralExtra }
  })

  const eocdZip64 = entries.length > U16_MAX || centralStart >= U32_MAX || centralTotal >= U32_MAX
  const out = new Uint8Array(localTotal + centralTotal + (eocdZip64 ? 76 : 0) + 22)
  const view = new DataView(out.buffer)
  let offset = 0

  for (const { entry, name, sizeZip64 } of resolved) {
    view.setUint32(offset, LOCAL_SIG, true)
    view.setUint16(offset + 4, sizeZip64 ? VERSION_ZIP64 : VERSION, true)
    view.setUint16(offset + 8, entry.method, true)
    view.setUint16(offset + 10, DOS_TIME, true)
    view.setUint16(offset + 12, DOS_DATE, true)
    view.setUint32(offset + 14, entry.crc, true)
    view.setUint32(offset + 18, sizeZip64 ? U32_MAX : entry.compressed.length, true)
    view.setUint32(offset + 22, sizeZip64 ? U32_MAX : entry.uncompressedSize, true)
    view.setUint16(offset + 26, name.length, true)
    view.setUint16(offset + 28, sizeZip64 ? 20 : 0, true)
    offset += 30
    out.set(name, offset)
    offset += name.length
    if (sizeZip64) {
      view.setUint16(offset, ZIP64_EXTRA_ID, true)
      view.setUint16(offset + 2, 16, true)
      setU64(view, offset + 4, entry.uncompressedSize)
      setU64(view, offset + 12, entry.compressed.length)
      offset += 20
    }
    out.set(entry.compressed, offset)
    offset += entry.compressed.length
  }

  for (const { entry, name, sizeZip64, offsetZip64, localOffset, centralExtra } of resolved) {
    view.setUint32(offset, CENTRAL_SIG, true)
    view.setUint16(offset + 4, VERSION, true)
    view.setUint16(offset + 6, sizeZip64 || offsetZip64 ? VERSION_ZIP64 : VERSION, true)
    view.setUint16(offset + 10, entry.method, true)
    view.setUint16(offset + 12, DOS_TIME, true)
    view.setUint16(offset + 14, DOS_DATE, true)
    view.setUint32(offset + 16, entry.crc, true)
    view.setUint32(offset + 20, sizeZip64 ? U32_MAX : entry.compressed.length, true)
    view.setUint32(offset + 24, sizeZip64 ? U32_MAX : entry.uncompressedSize, true)
    view.setUint16(offset + 28, name.length, true)
    view.setUint16(offset + 30, centralExtra, true)
    view.setUint32(offset + 42, offsetZip64 ? U32_MAX : localOffset, true)
    offset += 46
    out.set(name, offset)
    offset += name.length
    if (centralExtra > 0) {
      view.setUint16(offset, ZIP64_EXTRA_ID, true)
      view.setUint16(offset + 2, centralExtra - 4, true)
      offset += 4
      // The order is fixed: uncompressed, compressed, offset — only the fields
      // whose 32-bit slot was maxed, which is exactly what the reader expects.
      if (sizeZip64) {
        setU64(view, offset, entry.uncompressedSize)
        setU64(view, offset + 8, entry.compressed.length)
        offset += 16
      }
      if (offsetZip64) {
        setU64(view, offset, localOffset)
        offset += 8
      }
    }
  }

  if (eocdZip64) {
    const record = offset
    view.setUint32(offset, ZIP64_EOCD_SIG, true)
    setU64(view, offset + 4, 44) // size of the rest of this record
    view.setUint16(offset + 12, VERSION_ZIP64, true)
    view.setUint16(offset + 14, VERSION_ZIP64, true)
    setU64(view, offset + 24, entries.length)
    setU64(view, offset + 32, entries.length)
    setU64(view, offset + 40, centralTotal)
    setU64(view, offset + 48, centralStart)
    offset += 56
    view.setUint32(offset, ZIP64_LOCATOR_SIG, true)
    setU64(view, offset + 8, record)
    view.setUint32(offset + 16, 1, true) // total number of disks
    offset += 20
  }

  view.setUint32(offset, END_SIG, true)
  view.setUint16(offset + 8, Math.min(entries.length, U16_MAX), true)
  view.setUint16(offset + 10, Math.min(entries.length, U16_MAX), true)
  view.setUint32(offset + 12, Math.min(centralTotal, U32_MAX), true)
  view.setUint32(offset + 16, Math.min(centralStart, U32_MAX), true)
  return out
}
