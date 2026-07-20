import { type Container, writeContainer } from './container.js'
import { serialToDate } from './date.js'
import { type CellInput, patchSheet } from './patch.js'
import { type CellAddress, parseReference } from './reference.js'
import { type RawCell, readSheet } from './sheet.js'
import { readSharedStrings } from './shared-strings.js'
import { type Styles, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { type SheetState, readWorkbookPart } from './workbook.js'

export type CellValue =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: string }
  | { readonly kind: 'empty' }
  /** The stored number is kept so a date can be written back unchanged. */
  | { readonly kind: 'date'; readonly value: Date; readonly serial: number }

export interface Cell {
  readonly address: CellAddress
  readonly reference: string
  readonly value: CellValue
  readonly formula?: string
  /** The number format this cell resolves to, when it has one. */
  readonly numberFormat?: string
}

export interface Worksheet {
  readonly name: string
  readonly state: SheetState
  /** Streams the cells that carry content; empty cells are not visited. */
  cells(): Generator<Cell>
  /**
   * Records a new value for a cell. The change is visible to `cells()` and is
   * written by `toBytes()`; the rest of the sheet is left byte for byte alone.
   */
  set(reference: string, value: CellInput): void
}

export interface Workbook {
  readonly sheets: readonly Worksheet[]
  readonly date1904: boolean
  /**
   * Serialises the document back to .xlsx. Parts that were never interpreted
   * are written exactly as they were read.
   */
  toBytes(): Uint8Array
}

const EMPTY_STYLES: Styles = { numberFormats: new Map(), cellFormats: [] }

function partText(container: Container, path: string): string | undefined {
  const bytes = container.parts.get(path)
  return bytes === undefined ? undefined : new TextDecoder().decode(bytes)
}

function toCell(raw: RawCell, styles: Styles, date1904: boolean): Cell {
  const numberFormat = numberFormatOf(styles, raw.styleIndex)

  const value: CellValue =
    raw.value.kind === 'number' && isDateFormat(styles, raw.styleIndex)
      ? { kind: 'date', value: serialToDate(raw.value.value, date1904), serial: raw.value.value }
      : raw.value

  return {
    address: raw.address,
    reference: raw.reference,
    value,
    ...(raw.formula === undefined ? {} : { formula: raw.formula }),
    ...(numberFormat === undefined ? {} : { numberFormat }),
  }
}

export function readWorkbook(bytes: Uint8Array): Workbook {
  const part = readWorkbookPart(bytes)
  const { container, date1904 } = part

  const stylesXml = partText(container, 'xl/styles.xml')
  const styles = stylesXml === undefined ? EMPTY_STYLES : readStyles(stylesXml)

  const stringsXml = partText(container, 'xl/sharedStrings.xml')
  const sharedStrings = stringsXml === undefined ? [] : readSharedStrings(stringsXml)

  const edits = new Map<string, Map<string, CellInput>>()

  const sheets = part.sheets.map((reference): Worksheet => {
    const sheetXml = partText(container, reference.path)

    const patched = () => {
      if (sheetXml === undefined) return undefined
      const pending = edits.get(reference.path)
      return pending === undefined ? sheetXml : patchSheet(sheetXml, pending, date1904)
    }

    return {
      name: reference.name,
      state: reference.state,
      *cells(): Generator<Cell> {
        const xml = patched()
        if (xml === undefined) return
        for (const raw of readSheet(xml, sharedStrings)) {
          yield toCell(raw, styles, date1904)
        }
      },
      set(cellReference: string, value: CellInput): void {
        parseReference(cellReference)
        const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
        pending.set(cellReference, value)
        edits.set(reference.path, pending)
      },
    }
  })

  const toBytes = (): Uint8Array => {
    if (edits.size === 0) return writeContainer(container)

    const parts = new Map(container.parts)
    for (const [path, pending] of edits) {
      const xml = partText(container, path)
      if (xml === undefined) continue
      parts.set(path, new TextEncoder().encode(patchSheet(xml, pending, date1904)))
    }
    return writeContainer({ parts })
  }

  return { sheets, date1904, toBytes }
}
