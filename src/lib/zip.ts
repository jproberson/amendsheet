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
const ZIP64 = 0xffffffff

const notAZip = (message: string) => new XlsxError('not-a-zip', message, {})

function endOfCentralDirectory(view: DataView, length: number): number {
  // The record is 22 bytes plus an optional trailing comment, so it is found by
  // scanning back from the end for its signature.
  for (let offset = length - 22; offset >= 0; offset--) {
    if (view.getUint32(offset, true) === END_SIG) return offset
  }
  throw notAZip('File has no end-of-central-directory record, so it is not a zip')
}

/** Reads every entry's metadata and raw bytes, inflating nothing. */
export function readZip(bytes: Uint8Array): ZipEntry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = endOfCentralDirectory(view, bytes.length)

  const count = view.getUint16(end + 10, true)
  let offset = view.getUint32(end + 16, true)
  if (count === 0xffff || offset === ZIP64) {
    throw notAZip('Archive uses ZIP64 (over 4GB or 65535+ entries), which is not supported yet')
  }

  const entries: ZipEntry[] = []
  for (let index = 0; index < count; index++) {
    if (view.getUint32(offset, true) !== CENTRAL_SIG) {
      throw notAZip('Central directory is malformed')
    }
    const method = view.getUint16(offset + 10, true)
    const crc = view.getUint32(offset + 16, true)
    const compressedSize = view.getUint32(offset + 20, true)
    const uncompressedSize = view.getUint32(offset + 24, true)
    const nameLength = view.getUint16(offset + 28, true)
    const extraLength = view.getUint16(offset + 30, true)
    const commentLength = view.getUint16(offset + 32, true)
    const localOffset = view.getUint32(offset + 42, true)
    if (compressedSize === ZIP64 || uncompressedSize === ZIP64 || localOffset === ZIP64) {
      throw notAZip('Archive uses ZIP64, which is not supported yet')
    }

    const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength))

    if (view.getUint32(localOffset, true) !== LOCAL_SIG) {
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
    try {
      return inflateSync(entry.compressed, { out: new Uint8Array(entry.uncompressedSize) })
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

/** Writes the archive from entries that already carry their compressed bytes. */
export function writeZip(entries: readonly ZipEntry[]): Uint8Array {
  const encoder = new TextEncoder()
  const named = entries.map((entry) => ({ entry, name: encoder.encode(entry.name) }))

  let size = 22 // end-of-central-directory record
  for (const { entry, name } of named) {
    size += 30 + name.length + entry.compressed.length // local header + data
    size += 46 + name.length // central directory header
  }

  const out = new Uint8Array(size)
  const view = new DataView(out.buffer)
  let offset = 0
  const placed: Array<{ entry: ZipEntry; name: Uint8Array; localOffset: number }> = []

  for (const { entry, name } of named) {
    placed.push({ entry, name, localOffset: offset })
    view.setUint32(offset, LOCAL_SIG, true)
    view.setUint16(offset + 4, 20, true)
    view.setUint16(offset + 8, entry.method, true)
    view.setUint16(offset + 10, DOS_TIME, true)
    view.setUint16(offset + 12, DOS_DATE, true)
    view.setUint32(offset + 14, entry.crc, true)
    view.setUint32(offset + 18, entry.compressed.length, true)
    view.setUint32(offset + 22, entry.uncompressedSize, true)
    view.setUint16(offset + 26, name.length, true)
    offset += 30
    out.set(name, offset)
    offset += name.length
    out.set(entry.compressed, offset)
    offset += entry.compressed.length
  }

  const centralStart = offset
  for (const { entry, name, localOffset } of placed) {
    view.setUint32(offset, CENTRAL_SIG, true)
    view.setUint16(offset + 4, 20, true)
    view.setUint16(offset + 6, 20, true)
    view.setUint16(offset + 10, entry.method, true)
    view.setUint16(offset + 12, DOS_TIME, true)
    view.setUint16(offset + 14, DOS_DATE, true)
    view.setUint32(offset + 16, entry.crc, true)
    view.setUint32(offset + 20, entry.compressed.length, true)
    view.setUint32(offset + 24, entry.uncompressedSize, true)
    view.setUint16(offset + 28, name.length, true)
    view.setUint32(offset + 42, localOffset, true)
    offset += 46
    out.set(name, offset)
    offset += name.length
  }

  view.setUint32(offset, END_SIG, true)
  view.setUint16(offset + 8, named.length, true)
  view.setUint16(offset + 10, named.length, true)
  view.setUint32(offset + 12, offset - centralStart, true)
  view.setUint32(offset + 16, centralStart, true)
  return out
}
