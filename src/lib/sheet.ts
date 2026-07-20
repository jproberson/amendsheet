import { XlsxError } from './errors.js'
import { type CellAddress, formatReference, parseReference } from './reference.js'
import { readXml } from './xml.js'

export type RawCellValue =
  | { readonly kind: 'number'; readonly value: number }
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
  /** Index into the style table. Resolving it to a format needs styles.xml. */
  readonly styleIndex?: number
}

/**
 * Dates are deliberately not produced here. A date is a number wearing a date
 * number format, so telling them apart needs the style table; this reader
 * reports what the cell actually stores.
 */
export function* readSheet(xml: string, sharedStrings: readonly string[]): Generator<RawCell> {
  let row = 0
  let column = 0

  let type = ''
  let styleIndex: number | undefined
  let reference: string | undefined

  let rawValue: string[] | null = null
  let formula: string[] | null = null
  let inlineText: string[] | null = null

  let inValue = false
  let inFormula = false
  let inInlineText = false

  const finishCell = (): RawCell => {
    const address = reference === undefined ? { row, column } : parseReference(reference)
    const written = reference ?? formatReference(address)
    column = address.column

    return {
      address,
      reference: written,
      value: toValue(type, rawValue, inlineText, sharedStrings, written),
      ...(formula === null ? {} : { formula: formula.join('') }),
      ...(styleIndex === undefined ? {} : { styleIndex }),
    }
  }

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      switch (event.name) {
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
          inlineText = null
          column++
          // A self closing cell carries only a style, and gets no close event.
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
          break
        case 't':
          inlineText ??= []
          inInlineText = !event.selfClosing
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

    switch (event.name) {
      case 'v':
        inValue = false
        break
      case 'f':
        inFormula = false
        break
      case 't':
        inInlineText = false
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
): RawCellValue {
  if (type === 'inlineStr') {
    return { kind: 'text', value: inlineText === null ? '' : inlineText.join('') }
  }
  if (rawValue === null) return { kind: 'empty' }

  const raw = rawValue.join('')

  switch (type) {
    case 's': {
      const index = Number(raw)
      return { kind: 'text', value: sharedStrings[index] ?? '' }
    }
    case 'str':
      return { kind: 'text', value: raw }
    case 'b':
      return { kind: 'boolean', value: raw === '1' }
    case 'e':
      return { kind: 'error', value: raw }
    default: {
      if (raw === '') return { kind: 'empty' }
      const value = Number(raw)
      if (Number.isNaN(value)) {
        throw new XlsxError(`Cell ${reference} holds "${raw}", which is not a number`)
      }
      return { kind: 'number', value }
    }
  }
}
