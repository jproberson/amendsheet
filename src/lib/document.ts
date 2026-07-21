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
import { ensureDateStyle, ensureNumberFormat } from './styles-writer.js'
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
  const formats = new Map<string, Map<string, string>>()

  const sheets = part.sheets.map((reference): Worksheet => {
    const sheetXml = partText(container, reference.path)

    const patched = () => {
      if (sheetXml === undefined) return undefined
      const pending = edits.get(reference.path)
      return pending === undefined ? sheetXml : patchSheet(sheetXml, pending, date1904)
    }

    function* readCells(): Generator<Cell> {
      const xml = patched()
      if (xml === undefined) return
      for (const raw of readSheet(xml, sharedStrings)) {
        yield toCell(raw, styles, date1904)
      }
    }

    // Built once per edit, so reading many cells does not reparse the sheet
    // once per lookup.
    let byReference: Map<string, Cell> | undefined

    return {
      name: reference.name,
      state: reference.state,
      cells: readCells,
      cell(cellReference: string): Cell | undefined {
        if (byReference === undefined) {
          byReference = new Map()
          for (const found of readCells()) byReference.set(found.reference, found)
        }
        return byReference.get(formatReference(parseReference(cellReference)))
      },
      set(cellReference: string, value: CellInput, options?: WriteOptions): void {
        byReference = undefined
        // Normalised so `a1`, `$A$1` and `A1` are one edit, and so the file
        // never receives a reference spelled the way the caller typed it.
        const canonical = formatReference(parseWritableReference(cellReference))
        const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
        pending.set(canonical, value)
        edits.set(reference.path, pending)

        if (options?.format !== undefined) {
          const wanted = formats.get(reference.path) ?? new Map<string, string>()
          wanted.set(canonical, options.format)
          formats.set(reference.path, wanted)
        }
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

    // A Date needs a cell format that displays dates, or it shows as a serial.
    let workingStyles = stylesXml
    const dateStyles = new Map<string, Map<string, number>>()

    if (workingStyles !== undefined) {
      for (const [path, pending] of edits) {
        const sheetXml = partText(container, path)
        if (sheetXml === undefined) continue

        const existing = new Map<string, number | undefined>()
        for (const raw of readSheet(sheetXml, [])) existing.set(raw.reference, raw.styleIndex)

        const wanted = formats.get(path)
        const overrides = new Map<string, number>()

        for (const [reference, value] of pending) {
          const format = wanted?.get(reference)
          const current = existing.get(reference)

          // An asked-for format wins; a Date only gets one because it needs one.
          let applied: { xml: string; index: number } | undefined
          if (format !== undefined) applied = ensureNumberFormat(workingStyles, current, format)
          else if (value instanceof Date) applied = ensureDateStyle(workingStyles, current)
          if (applied === undefined) continue

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
