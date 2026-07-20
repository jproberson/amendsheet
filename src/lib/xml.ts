import { XlsxError } from './errors.js'

export type XmlEvent =
  | { kind: 'open'; name: string; attributes: ReadonlyMap<string, string>; selfClosing: boolean }
  | { kind: 'text'; text: string }
  | { kind: 'close'; name: string }

const ENTITY = /&(?:#x([0-9a-fA-F]+)|#(\d+)|([a-zA-Z][a-zA-Z0-9]*));/g

const NAMED_ENTITIES = new Map([
  ['amp', '&'],
  ['lt', '<'],
  ['gt', '>'],
  ['quot', '"'],
  ['apos', "'"],
])

function decodeEntities(text: string): string {
  if (!text.includes('&')) return text
  return text.replace(
    ENTITY,
    (match: string, hex?: string, decimal?: string, named?: string): string => {
      if (hex !== undefined) return String.fromCodePoint(Number.parseInt(hex, 16))
      if (decimal !== undefined) return String.fromCodePoint(Number(decimal))
      if (named !== undefined) return NAMED_ENTITIES.get(named) ?? match
      return match
    },
  )
}

const isWhitespace = (character: string) =>
  character === ' ' || character === '\t' || character === '\n' || character === '\r'

/** Finds the `>` that ends a tag, ignoring any that appear inside quoted values. */
function findTagEnd(source: string, start: number): number {
  let quote: string | null = null
  for (let index = start; index < source.length; index++) {
    const character = source.charAt(index)
    if (quote !== null) {
      if (character === quote) quote = null
    } else if (character === '"' || character === "'") {
      quote = character
    } else if (character === '>') {
      return index
    }
  }
  return -1
}

function parseAttributes(source: string, tagStart: number): ReadonlyMap<string, string> {
  const attributes = new Map<string, string>()
  let index = 0

  while (index < source.length) {
    while (index < source.length && isWhitespace(source.charAt(index))) index++
    if (index >= source.length) break

    const nameStart = index
    while (
      index < source.length &&
      !isWhitespace(source.charAt(index)) &&
      source.charAt(index) !== '='
    ) {
      index++
    }
    const name = source.slice(nameStart, index)
    if (name.length === 0) break

    while (index < source.length && isWhitespace(source.charAt(index))) index++
    if (source.charAt(index) !== '=') {
      throw new XlsxError(`Attribute "${name}" has no value, at offset ${tagStart}`)
    }
    index++

    while (index < source.length && isWhitespace(source.charAt(index))) index++
    const quote = source.charAt(index)
    if (quote !== '"' && quote !== "'") {
      throw new XlsxError(`Attribute "${name}" is not quoted, at offset ${tagStart}`)
    }
    index++

    // findTagEnd only stops outside quotes, so the closing quote is always present.
    const valueStart = index
    while (index < source.length && source.charAt(index) !== quote) index++
    attributes.set(name, decodeEntities(source.slice(valueStart, index)))
    index++
  }

  return attributes
}

/**
 * Emits elements and text as they are encountered rather than building a tree,
 * so a sheet with hundreds of thousands of rows can be read without holding all
 * of it in memory.
 */
export function* readXml(source: string): Generator<XmlEvent> {
  let position = 0

  while (position < source.length) {
    const tagStart = source.indexOf('<', position)

    if (tagStart === -1) {
      yield { kind: 'text', text: decodeEntities(source.slice(position)) }
      return
    }
    if (tagStart > position) {
      yield { kind: 'text', text: decodeEntities(source.slice(position, tagStart)) }
    }
    position = tagStart

    if (source.startsWith('<![CDATA[', position)) {
      const end = source.indexOf(']]>', position)
      if (end === -1) throw new XlsxError(`Unterminated CDATA at offset ${position}`)
      yield { kind: 'text', text: source.slice(position + 9, end) }
      position = end + 3
      continue
    }
    if (source.startsWith('<!--', position)) {
      const end = source.indexOf('-->', position)
      if (end === -1) throw new XlsxError(`Unterminated comment at offset ${position}`)
      position = end + 3
      continue
    }
    if (source.startsWith('<?', position)) {
      const end = source.indexOf('?>', position)
      if (end === -1)
        throw new XlsxError(`Unterminated processing instruction at offset ${position}`)
      position = end + 2
      continue
    }
    if (source.startsWith('<!', position)) {
      const end = findTagEnd(source, position)
      if (end === -1) throw new XlsxError(`Unterminated declaration at offset ${position}`)
      position = end + 1
      continue
    }

    const tagEnd = findTagEnd(source, position)
    if (tagEnd === -1) throw new XlsxError(`Unclosed tag starting at offset ${position}`)

    const inner = source.slice(position + 1, tagEnd)
    position = tagEnd + 1

    if (inner.startsWith('/')) {
      yield { kind: 'close', name: inner.slice(1).trim() }
      continue
    }

    const selfClosing = inner.endsWith('/')
    const body = selfClosing ? inner.slice(0, -1) : inner

    let nameEnd = 0
    while (nameEnd < body.length && !isWhitespace(body.charAt(nameEnd))) nameEnd++

    yield {
      kind: 'open',
      name: body.slice(0, nameEnd),
      attributes: parseAttributes(body.slice(nameEnd), tagStart),
      selfClosing,
    }
  }
}
