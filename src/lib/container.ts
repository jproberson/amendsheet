import { XlsxError } from './errors.js'
import { type ZipEntry, deflate, inflate, readZip, writeZip } from './zip.js'

/** Every part of the file, including the ones nothing here interprets. */
export interface Container {
  readonly parts: ReadonlyMap<string, Uint8Array>
}

/**
 * A container backed by its original bytes. `write` rebuilds the package,
 * compressing only the parts that changed and copying the rest through still
 * compressed, so a part nothing read is never inflated.
 */
export interface ReadableContainer extends Container {
  write(changes: ReadonlyMap<string, Uint8Array | null>): Uint8Array
}

/**
 * Inflates a part on first access and remembers it. A workbook edited at one
 * cell only ever inflates the handful of parts it reads; the rest stay
 * compressed in memory until they are copied, still compressed, into the output.
 */
class LazyParts implements ReadonlyMap<string, Uint8Array> {
  private readonly cache = new Map<string, Uint8Array>()
  private everything: Map<string, Uint8Array> | undefined

  constructor(private readonly raw: ReadonlyMap<string, ZipEntry>) {}

  private read(name: string, entry: ZipEntry): Uint8Array {
    const cached = this.cache.get(name)
    if (cached !== undefined) return cached
    const bytes = inflate(entry)
    this.cache.set(name, bytes)
    return bytes
  }

  get(name: string): Uint8Array | undefined {
    const entry = this.raw.get(name)
    return entry === undefined ? undefined : this.read(name, entry)
  }

  has(name: string): boolean {
    return this.raw.has(name)
  }

  get size(): number {
    return this.raw.size
  }

  keys(): MapIterator<string> {
    return this.raw.keys()
  }

  // Iterating values or pairs wants every part, so it inflates them all at once.
  private all(): Map<string, Uint8Array> {
    if (this.everything === undefined) {
      this.everything = new Map()
      for (const [name, entry] of this.raw) this.everything.set(name, this.read(name, entry))
    }
    return this.everything
  }

  values(): MapIterator<Uint8Array> {
    return this.all().values()
  }

  entries(): MapIterator<[string, Uint8Array]> {
    return this.all().entries()
  }

  [Symbol.iterator](): MapIterator<[string, Uint8Array]> {
    return this.all()[Symbol.iterator]()
  }

  forEach(
    callback: (value: Uint8Array, key: string, map: ReadonlyMap<string, Uint8Array>) => void,
    thisArg?: unknown,
  ): void {
    for (const [name, bytes] of this.all()) callback.call(thisArg, bytes, name, this)
  }
}

/** The magic of an OLE2 compound file: a legacy .xls, or an encrypted .xlsx. */
const OLE2_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]

const startsWith = (bytes: Uint8Array, magic: readonly number[]) =>
  magic.every((byte, index) => bytes[index] === byte)

export function readContainer(bytes: Uint8Array): ReadableContainer {
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

  const entries = new Map<string, ZipEntry>()
  for (const entry of readZip(bytes)) {
    // A directory entry is an artifact of how the zip was produced, not a part.
    if (entry.name.endsWith('/')) continue
    entries.set(entry.name, entry)
  }

  return {
    parts: new LazyParts(entries),
    write: (changes) => writeChanges(entries, changes),
  }
}

/**
 * Rebuilds the package: a part named in `changes` is deflated fresh (or dropped
 * when its change is null), and every other part is written with the compressed
 * bytes it was read with. Only a part the caller changed is ever recompressed.
 */
function writeChanges(
  entries: ReadonlyMap<string, ZipEntry>,
  changes: ReadonlyMap<string, Uint8Array | null>,
): Uint8Array {
  const out: ZipEntry[] = []
  for (const [name, entry] of entries) {
    const change = changes.get(name)
    if (change === undefined) out.push(entry)
    else if (change !== null) out.push(deflate(name, change))
  }
  return writeZip(out)
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

export function writeContainer(container: Container): Uint8Array {
  const entries: ZipEntry[] = []
  for (const [path, content] of container.parts) entries.push(deflate(path, content))
  return writeZip(entries)
}
