import { type Container, writeContainer } from './container.js'
import { serialToDate } from './date.js'
import type { CellAddress } from './reference.js'
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

  const sheets = part.sheets.map((reference): Worksheet => {
    const sheetXml = partText(container, reference.path)

    return {
      name: reference.name,
      state: reference.state,
      *cells(): Generator<Cell> {
        if (sheetXml === undefined) return
        for (const raw of readSheet(sheetXml, sharedStrings)) {
          yield toCell(raw, styles, date1904)
        }
      },
    }
  })

  return {
    sheets,
    date1904,
    toBytes: () => writeContainer(container),
  }
}
