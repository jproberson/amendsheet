import { type XlsxErrorContext, XlsxError } from './errors.js'
import { type CellAddress, formatReference, parseFileReference } from './reference.js'
import { readXmlBytes } from './xml.js'

export type RawCellValue =
  | { readonly kind: 'number'; readonly value: number }
  /** An ISO-8601 date written literally, which ECMA-376 permits. */
  | { readonly kind: 'date'; readonly value: string }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: string }
  | { readonly kind: 'empty' }

export interface RawCell {
  readonly address: CellAddress
  readonly reference: string
  readonly value: RawCellValue
  /** Formula source, without the leading `=`. Absent when the cell holds a literal. */
  readonly formula?: string
  /** The si of the shared group, on the master and on every dependent alike. */
  readonly sharedIndex?: string
  /** Set on the cell that owns the shared range, which is where the source is. */
  readonly ownsSharedRange?: boolean
  /** Resolving this to a format needs styles.xml. */
  readonly styleIndex?: number
}

/** No dates: a date is a number with a date format, which needs the style table. */
export function* readSheet(
  bytes: Uint8Array,
  sharedStrings: readonly string[],
  at: XlsxErrorContext = {},
): Generator<RawCell> {
  let row = 0
  let column = 0

  let type = ''
  let styleIndex: number | undefined
  let reference: string | undefined

  let rawValue: string[] | null = null
  let formula: string[] | null = null
  let sharedIndex: string | undefined
  let ownsSharedRange = false
  let inlineText: string[] | null = null

  let inValue = false
  let inFormula = false
  let inInlineText = false
  let inPhonetic = false

  const finishCell = (): RawCell => {
    const address = reference === undefined ? { row, column } : parseFileReference(reference, at)
    const written = reference ?? formatReference(address)
    column = address.column

    return {
      address,
      reference: written,
      value: toValue(type, rawValue, inlineText, sharedStrings, written, at),
      ...(formula === null ? {} : { formula: formula.join('') }),
      ...(sharedIndex === undefined ? {} : { sharedIndex, ownsSharedRange }),
      ...(styleIndex === undefined ? {} : { styleIndex }),
    }
  }

  for (const event of readXmlBytes(bytes)) {
    if (event.kind === 'open') {
      switch (event.localName) {
        case 'row': {
          const declared = event.attributes.get('r')
          row = declared === undefined ? row + 1 : Number(declared)
          column = 0
          break
        }
        case 'c': {
          reference = event.attributes.get('r')
          type = event.attributes.get('t') ?? 'n'
          const style = event.attributes.get('s')
          styleIndex = style === undefined ? undefined : Number(style)
          rawValue = null
          formula = null
          sharedIndex = undefined
          ownsSharedRange = false
          inlineText = null
          inPhonetic = false
          column++
          // A self closing cell gets no close event.
          if (event.selfClosing) yield finishCell()
          break
        }
        case 'v':
          rawValue = []
          inValue = !event.selfClosing
          break
        case 'f':
          formula = []
          inFormula = !event.selfClosing
          // Only the cell carrying ref owns the range; a dependent names the
          // si alone, and either may be written self closing.
          if (event.attributes.get('t') === 'shared') {
            sharedIndex = event.attributes.get('si')
            ownsSharedRange = event.attributes.get('ref') !== undefined
          }
          break
        case 'rPh':
          inPhonetic = true
          break
        case 't':
          inlineText ??= []
          inInlineText = !event.selfClosing && !inPhonetic
          break
      }
      continue
    }

    if (event.kind === 'text') {
      if (inValue && rawValue !== null) rawValue.push(event.text)
      else if (inFormula && formula !== null) formula.push(event.text)
      else if (inInlineText && inlineText !== null) inlineText.push(event.text)
      continue
    }

    switch (event.localName) {
      case 'v':
        inValue = false
        break
      case 'f':
        inFormula = false
        break
      case 't':
        inInlineText = false
        break
      case 'rPh':
        inPhonetic = false
        break
      case 'c':
        yield finishCell()
        break
    }
  }
}

function toValue(
  type: string,
  rawValue: string[] | null,
  inlineText: string[] | null,
  sharedStrings: readonly string[],
  reference: string,
  at: XlsxErrorContext,
): RawCellValue {
  if (type === 'inlineStr') {
    return { kind: 'text', value: inlineText === null ? '' : inlineText.join('') }
  }
  if (rawValue === null) return { kind: 'empty' }

  const raw = rawValue.join('')

  switch (type) {
    case 's': {
      const value = sharedStrings[Number(raw)]
      if (value === undefined) {
        throw new XlsxError(
          'invalid-content',
          `Cell ${reference} references shared string ${raw}, which the table does not hold`,
          { ...at, reference },
        )
      }
      return { kind: 'text', value }
    }
    case 'str':
      return { kind: 'text', value: raw }
    case 'd':
      return { kind: 'date', value: raw }
    case 'b':
      return { kind: 'boolean', value: raw === '1' }
    case 'e':
      return { kind: 'error', value: raw }
    default: {
      if (raw === '') return { kind: 'empty' }
      const value = Number(raw)
      if (Number.isNaN(value)) {
        throw new XlsxError(
          'invalid-content',
          `Cell ${reference} holds "${raw}", which is not a number`,
          { ...at, reference },
        )
      }
      return { kind: 'number', value }
    }
  }
}
