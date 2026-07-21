import { XlsxError } from './errors.js'
import { bumpAttribute, findUnwritableCharacter, readXml } from './xml.js'

/**
 * One string may be split across formatting runs, and `rPh` runs hold phonetic
 * guides that look like text but are not part of the value.
 */
interface SharedString {
  readonly text: string
  /**
   * A single unformatted `t`. An entry built from runs, or carrying a phonetic
   * guide, reads as the same text but carries formatting with it, so a write
   * must not be pointed at one.
   */
  readonly plain: boolean
}

function readEntries(xml: string): readonly SharedString[] {
  const entries: SharedString[] = []

  let current: string[] | null = null
  let plain = true
  let inPhonetic = false
  let inText = false

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.localName === 'si') {
        current = []
        plain = true
      } else if (event.localName === 'rPh') {
        inPhonetic = true
        plain = false
      } else if (event.localName === 'r') plain = false
      else if (event.localName === 't' && !event.selfClosing) inText = !inPhonetic
      continue
    }

    if (event.kind === 'text') {
      if (inText && current !== null) current.push(event.text)
      continue
    }

    if (event.localName === 't') inText = false
    else if (event.localName === 'rPh') inPhonetic = false
    else if (event.localName === 'si' && current !== null) {
      entries.push({ text: current.join(''), plain })
      current = null
    }
  }

  return entries
}

export function readSharedStrings(xml: string): readonly string[] {
  return readEntries(xml).map((entry) => entry.text)
}

export interface AppendedStrings {
  readonly xml: string
  /** Index in the table for every string that was asked for. */
  readonly indexes: ReadonlyMap<string, number>
}

const escapeXml = (text: string) =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // XML 1.0 has a parser fold a literal CR, and the CR of a CRLF, into a bare
    // LF before the document reaches the application. Written raw it is lost.
    .replace(/\r/g, '&#13;')

function entryFor(value: string, prefix: string): string {
  const unwritable = findUnwritableCharacter(value)
  if (unwritable !== undefined) {
    throw new XlsxError(
      'unwritable-value',
      `A string holds ${unwritable}, which cannot be written to xml`,
    )
  }
  const needsPreserve = value !== value.trim()
  const attributes = needsPreserve ? ' xml:space="preserve"' : ''
  return `<${prefix}si><${prefix}t${attributes}>${escapeXml(value)}</${prefix}t></${prefix}si>`
}

/**
 * Adds strings the table does not have yet and reports where every requested
 * string lives. Existing entries are copied through untouched.
 */
export function appendSharedStrings(xml: string, strings: readonly string[]): AppendedStrings {
  const existing = readEntries(xml)
  const indexes = new Map<string, number>()
  for (const [index, entry] of existing.entries()) {
    if (entry.plain && !indexes.has(entry.text)) indexes.set(entry.text, index)
  }

  let openTag = ''
  let insertAt = -1
  let closeLength = 0
  // Some writers prefix every element. An entry written without the prefix
  // lands in a namespace the document may not even bind, so the strings are
  // there but no cell resolves to them.
  let prefix = ''

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'sst') {
      openTag = xml.slice(event.start, event.end)
      const colon = event.name.indexOf(':')
      prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
      if (event.selfClosing) {
        insertAt = event.end
        closeLength = event.end - event.start
      }
      continue
    }
    if (event.kind === 'close' && event.localName === 'sst') {
      insertAt = event.start
      closeLength = 0
    }
  }

  const additions: string[] = []
  const requested = new Map<string, number>()

  for (const value of strings) {
    const known = indexes.get(value)
    if (known !== undefined) {
      requested.set(value, known)
      continue
    }
    const index = existing.length + additions.length
    indexes.set(value, index)
    requested.set(value, index)
    additions.push(entryFor(value, prefix))
  }

  if (additions.length === 0) return { xml, indexes: requested }

  if (insertAt === -1)
    throw new XlsxError('malformed-xml', 'Shared string table has no sst element')

  const updatedTag = bumpAttribute(
    bumpAttribute(openTag, 'uniqueCount', additions.length),
    'count',
    additions.length,
  )

  const body = additions.join('')

  if (closeLength > 0) {
    const opened = `${updatedTag.slice(0, -2)}>`
    return {
      xml: `${xml.slice(0, insertAt - closeLength)}${opened}${body}</${prefix}sst>${xml.slice(insertAt)}`,
      indexes: requested,
    }
  }

  const head = xml.slice(0, insertAt).replace(openTag, updatedTag)
  return { xml: `${head}${body}${xml.slice(insertAt)}`, indexes: requested }
}
