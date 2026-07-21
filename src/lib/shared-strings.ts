import { XlsxError } from './errors.js'
import { findUnwritableCharacter, readXml } from './xml.js'

/**
 * One string may be split across formatting runs, and `rPh` runs hold phonetic
 * guides that look like text but are not part of the value.
 */
export function readSharedStrings(xml: string): readonly string[] {
  const strings: string[] = []

  let current: string[] | null = null
  let inPhonetic = false
  let inText = false

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.localName === 'si') current = []
      else if (event.localName === 'rPh') inPhonetic = true
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
      strings.push(current.join(''))
      current = null
    }
  }

  return strings
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

function bumpAttribute(openTag: string, name: string, by: number): string {
  return openTag.replace(new RegExp(`${name}="(\\d+)"`), (match, digits: string) =>
    match.replace(digits, String(Number(digits) + by)),
  )
}

/**
 * Adds strings the table does not have yet and reports where every requested
 * string lives. Existing entries are copied through untouched.
 */
export function appendSharedStrings(xml: string, strings: readonly string[]): AppendedStrings {
  const existing = readSharedStrings(xml)
  const indexes = new Map<string, number>()
  for (const [index, value] of existing.entries()) {
    if (!indexes.has(value)) indexes.set(value, index)
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
