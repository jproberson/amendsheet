import { XlsxError } from './errors.js'

/** Byte range in the source, so the original can be spliced exactly. */
interface Span {
  readonly start: number
  readonly end: number
}

export type XmlEvent =
  | (Span & {
      kind: 'open'
      /** As written, prefix included, so the original can be reproduced. */
      name: string
      /** Without the namespace prefix, which is the file's choice. */
      localName: string
      attributes: ReadonlyMap<string, string>
      selfClosing: boolean
    })
  | (Span & { kind: 'text'; text: string })
  | (Span & { kind: 'close'; name: string; localName: string })

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
      throw new XlsxError(
        'malformed-xml',
        `Attribute "${name}" has no value, at offset ${tagStart}`,
      )
    }
    index++

    while (index < source.length && isWhitespace(source.charAt(index))) index++
    const quote = source.charAt(index)
    if (quote !== '"' && quote !== "'") {
      throw new XlsxError(
        'malformed-xml',
        `Attribute "${name}" is not quoted, at offset ${tagStart}`,
      )
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

const localNameOf = (name: string) => {
  const colon = name.indexOf(':')
  return colon === -1 ? name : name.slice(colon + 1)
}

/** Emits events rather than building a tree, so large sheets stream. */
export function* readXml(source: string): Generator<XmlEvent> {
  let position = 0

  while (position < source.length) {
    const tagStart = source.indexOf('<', position)

    if (tagStart === -1) {
      yield {
        kind: 'text',
        text: decodeEntities(source.slice(position)),
        start: position,
        end: source.length,
      }
      return
    }
    if (tagStart > position) {
      yield {
        kind: 'text',
        text: decodeEntities(source.slice(position, tagStart)),
        start: position,
        end: tagStart,
      }
    }
    position = tagStart

    if (source.startsWith('<![CDATA[', position)) {
      const end = source.indexOf(']]>', position)
      if (end === -1)
        throw new XlsxError('malformed-xml', `Unterminated CDATA at offset ${position}`)
      yield { kind: 'text', text: source.slice(position + 9, end), start: position, end: end + 3 }
      position = end + 3
      continue
    }
    if (source.startsWith('<!--', position)) {
      const end = source.indexOf('-->', position)
      if (end === -1)
        throw new XlsxError('malformed-xml', `Unterminated comment at offset ${position}`)
      position = end + 3
      continue
    }
    if (source.startsWith('<?', position)) {
      const end = source.indexOf('?>', position)
      if (end === -1)
        throw new XlsxError(
          'malformed-xml',
          `Unterminated processing instruction at offset ${position}`,
        )
      position = end + 2
      continue
    }
    if (source.startsWith('<!', position)) {
      const end = findTagEnd(source, position)
      if (end === -1)
        throw new XlsxError('malformed-xml', `Unterminated declaration at offset ${position}`)
      position = end + 1
      continue
    }

    const tagEnd = findTagEnd(source, position)
    if (tagEnd === -1)
      throw new XlsxError('malformed-xml', `Unclosed tag starting at offset ${position}`)

    const inner = source.slice(position + 1, tagEnd)
    position = tagEnd + 1

    if (inner.startsWith('/')) {
      const name = inner.slice(1).trim()
      yield { kind: 'close', name, localName: localNameOf(name), start: tagStart, end: position }
      continue
    }

    const selfClosing = inner.endsWith('/')
    const body = selfClosing ? inner.slice(0, -1) : inner

    let nameEnd = 0
    while (nameEnd < body.length && !isWhitespace(body.charAt(nameEnd))) nameEnd++

    const name = body.slice(0, nameEnd)
    yield {
      kind: 'open',
      name,
      localName: localNameOf(name),
      attributes: parseAttributes(body.slice(nameEnd), tagStart),
      selfClosing,
      start: tagStart,
      end: position,
    }
  }
}

const name = (code: number) => `U+${code.toString(16).toUpperCase().padStart(4, '0')}`

/** Returns the offending code point, or undefined when the text is writable. */
export function findUnwritableCharacter(text: string): string | undefined {
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index)

    // XML 1.0 permits only tab, newline and carriage return below U+0020.
    if (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) return name(code)

    // Not characters at all, and outside XML's Char production.
    if (code === 0xfffe || code === 0xffff) return name(code)

    // A pair is one character above the basic plane and perfectly writable; a
    // half on its own is not, and encoding replaces it with U+FFFD, so the
    // value in hand and the value in the file would quietly differ.
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(index + 1)
      if (Number.isNaN(next) || next < 0xdc00 || next > 0xdfff) return name(code)
      index++
      continue
    }
    if (code >= 0xdc00 && code <= 0xdfff) return name(code)
  }
  return undefined
}
