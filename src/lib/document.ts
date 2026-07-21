import { type Container, writeContainer } from './container.js'
import { serialToDate } from './date.js'
import { type CellInput, patchSheet } from './patch.js'
import { type CellAddress, formatReference, parseReference } from './reference.js'
import { type RawCell, readSheet } from './sheet.js'
import { appendSharedStrings, readSharedStrings } from './shared-strings.js'
import { ensureDateStyle } from './styles-writer.js'
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
  readonly numberFormat?: string
}

export interface Worksheet {
  readonly name: string
  readonly state: SheetState
  /**
   * Every cell the sheet stores. A cell that was cleared, or that carries only
   * formatting, is still stored, and arrives with a value of `kind: 'empty'`.
   */
  cells(): Generator<Cell>
  /** Undefined when the sheet stores nothing at that reference. */
  cell(reference: string): Cell | undefined
  /** Visible to `cells()` and `cell()` immediately, written by `toBytes()`. */
  set(reference: string, value: CellInput): void
}

export interface Workbook {
  readonly sheets: readonly Worksheet[]
  /** Undefined when no sheet has that name. Names are compared exactly. */
  sheet(name: string): Worksheet | undefined
  readonly date1904: boolean
  /** Parts that were never interpreted are written exactly as they were read. */
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
      cell(cellReference: string): Cell | undefined {
        const wanted = formatReference(parseReference(cellReference))
        for (const cell of this.cells()) {
          if (cell.reference === wanted) return cell
        }
        return undefined
      },
      set(cellReference: string, value: CellInput): void {
        // Normalised so `a1`, `$A$1` and `A1` are one edit, and so the file
        // never receives a reference spelled the way the caller typed it.
        const canonical = formatReference(parseReference(cellReference))
        const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
        pending.set(canonical, value)
        edits.set(reference.path, pending)
      },
    }
  })

  const toBytes = (): Uint8Array => {
    if (edits.size === 0) return writeContainer(container)

    const parts = new Map(container.parts)
    const encoder = new TextEncoder()

    // A Date needs a cell format that displays dates, or it shows as a serial.
    let workingStyles = stylesXml
    const dateStyles = new Map<string, Map<string, number>>()

    if (workingStyles !== undefined) {
      for (const [path, pending] of edits) {
        const sheetXml = partText(container, path)
        if (sheetXml === undefined) continue

        const existing = new Map<string, number | undefined>()
        for (const raw of readSheet(sheetXml, [])) existing.set(raw.reference, raw.styleIndex)

        const overrides = new Map<string, number>()
        for (const [reference, value] of pending) {
          if (!(value instanceof Date)) continue
          const current = existing.get(reference)
          const applied = ensureDateStyle(workingStyles, current)
          workingStyles = applied.xml
          if (applied.index !== current) overrides.set(reference, applied.index)
        }
        if (overrides.size > 0) dateStyles.set(path, overrides)
      }

      if (workingStyles !== stylesXml) {
        parts.set('xl/styles.xml', encoder.encode(workingStyles))
      }
    }

    // Text goes into the shared string table when the file has one, so the same
    // words written into many cells are stored once.
    let indexes: ReadonlyMap<string, number> | undefined
    if (stringsXml !== undefined) {
      const written: string[] = []
      for (const pending of edits.values()) {
        for (const value of pending.values()) {
          if (typeof value === 'string') written.push(value)
        }
      }
      if (written.length > 0) {
        const appended = appendSharedStrings(stringsXml, written)
        parts.set('xl/sharedStrings.xml', encoder.encode(appended.xml))
        indexes = appended.indexes
      }
    }

    for (const [path, pending] of edits) {
      const xml = partText(container, path)
      if (xml === undefined) continue
      parts.set(
        path,
        encoder.encode(patchSheet(xml, pending, date1904, indexes, dateStyles.get(path))),
      )
    }
    return writeContainer({ parts })
  }

  return {
    sheets,
    sheet: (name: string) => sheets.find((candidate) => candidate.name === name),
    date1904,
    toBytes,
  }
}
