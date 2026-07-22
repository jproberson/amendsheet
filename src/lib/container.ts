import { type Zippable, unzipSync, zipSync } from 'fflate'
import { XlsxError } from './errors.js'

/** Every part of the file, including the ones nothing here interprets. */
export interface Container {
  readonly parts: ReadonlyMap<string, Uint8Array>
}

/** The magic of an OLE2 compound file: a legacy .xls, or an encrypted .xlsx. */
const OLE2_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

const startsWith = (bytes: Uint8Array, magic: readonly number[]) =>
  magic.every((byte, index) => bytes[index] === byte)

export function readContainer(bytes: Uint8Array): Container {
  // An .xlsx is a zip. An OLE2 file reaching here is a password-protected
  // workbook or a legacy .xls, and "not a zip" points the user the wrong way.
  if (startsWith(bytes, OLE2_MAGIC)) {
    throw new XlsxError(
      'not-a-zip',
      'File is an OLE2 compound file, not an .xlsx. A password-protected workbook ' +
        'must be decrypted first, and a legacy .xls must be converted to .xlsx.',
      {},
    )
  }

  let entries: Record<string, Uint8Array>
  try {
    entries = unzipSync(bytes)
  } catch (cause) {
    throw new XlsxError('not-a-zip', 'File is not a zip archive, so it cannot be an .xlsx file', {
      cause,
    })
  }

  const parts = new Map<string, Uint8Array>()
  for (const [path, content] of Object.entries(entries)) {
    if (path.endsWith('/')) continue
    parts.set(path, content)
  }

  return { parts }
}

/**
 * Decodes an XML part as UTF-8, which is what OOXML parts are. A UTF-16 part is
 * caught rather than let through: XML forbids U+0000 and a UTF-8 part never
 * opens with 0xFF or 0xFE, so either at the start means UTF-16. Left to the
 * decoder, a UTF-16 part with a byte-order mark fails on the first byte, but one
 * without decodes to interleaved-null garbage that yields an empty part and no
 * error at all.
 */
export function decodeXmlPart(bytes: Uint8Array, path: string): string {
  const misencoded =
    bytes[0] === 0xff || bytes[0] === 0xfe || bytes[0] === 0x00 || bytes[1] === 0x00
  if (misencoded) {
    throw new XlsxError(
      'unreadable-part',
      `Part ${path} is not UTF-8 (it looks like UTF-16); every part must be UTF-8 XML`,
      { part: path },
    )
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch (cause) {
    throw new XlsxError('unreadable-part', `Part ${path} is not valid utf-8`, { part: path, cause })
  }
}

/** The ZIP epoch. A fixed stamp keeps the same parts producing the same bytes. */
const FIXED_TIMESTAMP = new Date(1980, 0, 1)

export function writeContainer(container: Container): Uint8Array {
  const entries: Zippable = {}
  for (const [path, content] of container.parts) {
    entries[path] = [content, { mtime: FIXED_TIMESTAMP }]
  }
  return zipSync(entries, { level: 6 })
}
