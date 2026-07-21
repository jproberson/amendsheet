import { type Container, writeContainer } from './container.js'
import { XlsxError } from './errors.js'
import { dateToSerial, serialToDate } from './date.js'
import { type CellInput, patchSheet } from './patch.js'
import {
  type CellAddress,
  formatReference,
  parseReference,
  parseWritableReference,
} from './reference.js'
import { type RawCell, readSheet } from './sheet.js'
import { appendSharedStrings, readSharedStrings } from './shared-strings.js'
import { type DateStyle, ensureDateStyle, ensureNumberFormat } from './styles-writer.js'
import { type Styles, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { readXml } from './xml.js'
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
  cells(): IterableIterator<Cell>
  /** Undefined when the sheet stores nothing at that reference. */
  cell(reference: string): Cell | undefined
  /**
   * Visible to `cells()` and `cell()` immediately, written by `toBytes()`.
   * A `format` is a number format code such as `"$"#,##0.00`; without one the
   * cell keeps the formatting it already had.
   */
  set(reference: string, value: CellInput, options?: WriteOptions): void
}

export interface WriteOptions {
  /** A number format code, applied to the cell being written. */
  readonly format?: string
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

const CALCULATION_CHAIN = 'xl/calcChain.xml'
const CONTENT_TYPES = '[Content_Types].xml'

/**
 * A written formula carries no computed result, so the workbook is marked for
 * recalculation. Without it a reader that trusts cached values shows nothing.
 */
function withRecalculation(xml: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'calcPr') continue

    const tag = xml.slice(event.start, event.end)
    if (tag.includes('fullCalcOnLoad=')) {
      return (
        xml.slice(0, event.start) +
        tag.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"') +
        xml.slice(event.end)
      )
    }
    const opened = tag.replace(/\/?>$/, (end) =>
      end === '/>' ? ' fullCalcOnLoad="1"/>' : ' fullCalcOnLoad="1">',
    )
    return xml.slice(0, event.start) + opened + xml.slice(event.end)
  }

  for (const event of readXml(xml)) {
    if (event.kind !== 'close' || event.localName !== 'workbook') continue
    const colon = event.name.indexOf(':')
    const prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
    const element = `<${prefix}calcPr fullCalcOnLoad="1"/>`
    return xml.slice(0, event.start) + element + xml.slice(event.start)
  }

  return xml
}

/** Removes one Override element, leaving every other byte of the part alone. */
function withoutOverride(xml: string, part: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'Override') continue
    if (event.attributes.get('PartName') !== `/${part}`) continue
    return xml.slice(0, event.start) + xml.slice(event.end)
  }
  return xml
}

function partText(container: Container, path: string): string | undefined {
  const bytes = container.parts.get(path)
  if (bytes === undefined) return undefined
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch (cause) {
    throw new XlsxError('unreadable-part', `Part ${path} is not valid utf-8`, { part: path, cause })
  }
}

function toCellValue(raw: RawCell, styles: Styles, date1904: boolean): CellValue {
  const value = raw.value

  if (value.kind === 'date') {
    const parsed = new Date(value.value)
    if (Number.isNaN(parsed.getTime())) return { kind: 'text', value: value.value }
    return { kind: 'date', value: parsed, serial: dateToSerial(parsed, date1904) }
  }

  if (value.kind === 'number' && isDateFormat(styles, raw.styleIndex)) {
    const serial = value.value
    // A serial outside the range dates cover stays the number it is.
    if (serial >= 0) {
      return { kind: 'date', value: serialToDate(serial, date1904), serial }
    }
  }

  return value
}

function toCell(raw: RawCell, styles: Styles, date1904: boolean): Cell {
  const numberFormat = numberFormatOf(styles, raw.styleIndex)
  const value = toCellValue(raw, styles, date1904)

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

  // Which cell format each edited cell lands on is decided by set(), not by
  // toBytes(). Choosing a style index is what makes a number a date, so it is a
  // decision about what a cell MEANS; leaving it to write time gave the read
  // path its own copy of the decision, and the two drifted.
  const styleOverrides = new Map<string, Map<string, number>>()
  let workingStyles = stylesXml
  let parsedStyles = styles
  let parsedFrom = stylesXml

  const stylesNow = (): Styles => {
    if (workingStyles !== undefined && workingStyles !== parsedFrom) {
      parsedStyles = readStyles(workingStyles)
      parsedFrom = workingStyles
    }
    return parsedStyles
  }

  const sheets = part.sheets.map((reference): Worksheet => {
    const sheetXml = partText(container, reference.path)

    const patched = () => {
      if (sheetXml === undefined) return undefined
      const pending = edits.get(reference.path)
      if (pending === undefined) return sheetXml
      return patchSheet(sheetXml, pending, date1904, undefined, styleOverrides.get(reference.path))
    }

    function* readCells(source?: string): Generator<Cell> {
      const xml = source ?? patched()
      if (xml === undefined) return
      for (const raw of readSheet(xml, sharedStrings)) yield toCell(raw, stylesNow(), date1904)
    }

    // The style each cell carried when the file was read, keyed canonically so
    // a file spelling a reference `a1` or `$A$1` still matches an edit to A1.
    let originalStyles: Map<string, number> | undefined
    const styleAt = (canonical: string): number | undefined => {
      const override = styleOverrides.get(reference.path)?.get(canonical)
      if (override !== undefined) return override

      if (originalStyles === undefined) {
        originalStyles = new Map()
        for (const raw of sheetXml === undefined ? [] : readSheet(sheetXml, [])) {
          if (raw.styleIndex !== undefined) {
            originalStyles.set(formatReference(raw.address), raw.styleIndex)
          }
        }
      }
      return originalStyles.get(canonical)
    }

    /**
     * What the cell becomes, resolved through the same function a read uses.
     * Predicting the value independently is how the overlay came to disagree
     * with the file it was about to write.
     */
    const predict = (canonical: string, value: CellInput, styleIndex: number | undefined): Cell => {
      const address = parseReference(canonical)
      const style = styleIndex === undefined ? {} : { styleIndex }
      const raw = { address, reference: canonical, ...style }

      if (value === null) {
        return toCell({ ...raw, value: { kind: 'empty' } }, stylesNow(), date1904)
      }
      if (typeof value === 'number') {
        return toCell({ ...raw, value: { kind: 'number', value } }, stylesNow(), date1904)
      }
      if (typeof value === 'boolean') {
        return toCell({ ...raw, value: { kind: 'boolean', value } }, stylesNow(), date1904)
      }
      if (typeof value === 'string') {
        return toCell({ ...raw, value: { kind: 'text', value } }, stylesNow(), date1904)
      }
      if (value instanceof Date) {
        // Written as the serial it becomes, so the style decides how it reads.
        const serial = dateToSerial(value, date1904)
        return toCell({ ...raw, value: { kind: 'number', value: serial } }, stylesNow(), date1904)
      }
      // A formula is written without a cached value, so it reads back empty.
      const formula = value.formula
      return toCell({ ...raw, value: { kind: 'empty' }, formula }, stylesNow(), date1904)
    }

    // Built from the sheet as it was read. Every edit is in the overlay, so the
    // index never needs rebuilding.
    let byReference: Map<string, Cell> | undefined
    const overlay = new Map<string, Cell>()

    return {
      name: reference.name,
      state: reference.state,
      cells: () => readCells(),
      cell(cellReference: string): Cell | undefined {
        const wanted = formatReference(parseReference(cellReference))
        const edited = overlay.get(wanted)
        if (edited !== undefined) return edited

        if (byReference === undefined) {
          byReference = new Map()
          for (const found of readCells(sheetXml)) {
            byReference.set(formatReference(found.address), found)
          }
        }
        return byReference.get(wanted)
      },
      set(cellReference: string, value: CellInput, options?: WriteOptions): void {
        // Normalised so `a1`, `$A$1` and `A1` are one edit, and so the file
        // never receives a reference spelled the way the caller typed it.
        const canonical = formatReference(parseWritableReference(cellReference))
        const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
        pending.set(canonical, value)
        edits.set(reference.path, pending)

        const current = styleAt(canonical)
        let resolved = current

        if (workingStyles !== undefined) {
          // An asked-for format wins; a Date only gets one because without one
          // it displays as the serial number it is stored as.
          let applied: DateStyle | undefined
          if (options?.format !== undefined) {
            applied = ensureNumberFormat(workingStyles, current, options.format)
          } else if (value instanceof Date) {
            applied = ensureDateStyle(workingStyles, current)
          }

          if (applied !== undefined) {
            workingStyles = applied.xml
            resolved = applied.index
            if (applied.index !== current) {
              const overrides = styleOverrides.get(reference.path) ?? new Map<string, number>()
              overrides.set(canonical, applied.index)
              styleOverrides.set(reference.path, overrides)
            }
          }
        }

        overlay.set(canonical, predict(canonical, value, resolved))
      },
    }
  })

  const toBytes = (): Uint8Array => {
    if (edits.size === 0) return writeContainer(container)

    const parts = new Map(container.parts)
    const encoder = new TextEncoder()

    // Excel rebuilds the calculation chain, but a stale one makes it offer to
    // repair the file, so the part and its content type go together.
    if (parts.delete(CALCULATION_CHAIN)) {
      const types = partText(container, CONTENT_TYPES)
      if (types !== undefined) {
        parts.set(CONTENT_TYPES, encoder.encode(withoutOverride(types, CALCULATION_CHAIN)))
      }
    }

    // set() already resolved every style; only the serialising is left.
    if (workingStyles !== undefined && workingStyles !== stylesXml) {
      parts.set('xl/styles.xml', encoder.encode(workingStyles))
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
        encoder.encode(patchSheet(xml, pending, date1904, indexes, styleOverrides.get(path))),
      )
    }

    const wroteFormula = [...edits.values()].some((pending) =>
      [...pending.values()].some(
        (value) => typeof value === 'object' && value !== null && !(value instanceof Date),
      ),
    )
    if (wroteFormula) {
      const book = partText(container, part.path)
      if (book !== undefined) parts.set(part.path, encoder.encode(withRecalculation(book)))
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
