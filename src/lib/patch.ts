import { dateToSerial } from './date.js'
import { XlsxError } from './errors.js'
import { formatReference, parseReference } from './reference.js'
import { findUnwritableCharacter, readXml } from './xml.js'

/** An expression without the leading `=`, so text starting with `=` stays text. */
export interface FormulaInput {
  readonly formula: string
}

export type CellInput = number | string | boolean | Date | null | FormulaInput

/** Element content only. Quotes need no escaping there, and Excel leaves them. */
const escapeXml = (text: string) =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // XML 1.0 has a parser fold a literal CR, and the CR of a CRLF, into a bare
    // LF before the document reaches the application. Written raw it is lost.
    .replace(/\r/g, '&#13;')

/**
 * A formula, or a refusal. Types stop a TypeScript caller passing anything
 * else; a JS caller, a JSON payload or an `any` at a boundary reach here, and
 * reading `.formula` off them threw outside the error contract.
 */
function formulaOf(value: object, reference: string): string {
  if ('formula' in value && typeof value.formula === 'string') return value.formula
  throw new XlsxError(
    'unwritable-value',
    `Cell ${reference} was given an object that names no string formula`,
    { reference },
  )
}

/**
 * Every refusal writing a cell can make, so `set()` can make it at the call
 * instead of leaving a workbook that only fails once it is saved. Takes
 * `unknown` because it is the boundary the value has to get past.
 */
export function checkWritable(reference: string, value: unknown, date1904: boolean): void {
  if (value === null || typeof value === 'boolean') return

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new XlsxError('unwritable-value', `Cell ${reference} cannot hold ${value}`, {
        reference,
      })
    }
    return
  }

  if (value instanceof Date) {
    dateToSerial(value, date1904, reference)
    return
  }

  if (typeof value !== 'string' && (typeof value !== 'object' || value === null)) {
    throw new XlsxError(
      'unwritable-value',
      `Cell ${reference} cannot hold a value of type ${typeof value}`,
      { reference },
    )
  }

  const unwritable = findUnwritableCharacter(
    typeof value === 'string' ? value : formulaOf(value, reference),
  )
  if (unwritable !== undefined) {
    throw new XlsxError(
      'unwritable-value',
      `Cell ${reference} holds ${unwritable}, which cannot be written to xml`,
      { reference },
    )
  }
}

/**
 * A formula whose result covers more cells than the one holding it. The others
 * carry a cached value and no formula of their own, so replacing this cell
 * leaves them owned by nothing.
 */
export interface SpillingFormula {
  readonly kind: 'shared' | 'array' | 'dataTable'
  /** The si for a shared formula, the covered range for the other two. */
  readonly name: string
}

/** Both keyed by canonical reference, so `a1` and `$A$1` find the same cell. */
export interface SheetIndex {
  /** The style index a cell carried when the file was read. */
  readonly styles: ReadonlyMap<string, number>
  /** Cells that other cells depend on for their formula. */
  readonly sharedFormulas: ReadonlyMap<string, SpillingFormula>
}

/** One pass, because set() needs both of these before it accepts an edit. */
export function indexSheet(xml: string): SheetIndex {
  const styles = new Map<string, number>()
  const sharedFormulas = new Map<string, SpillingFormula>()

  for (const row of readShape(xml).rows) {
    for (const cell of row.cells) {
      if (cell.style === undefined && cell.spillingFormula === undefined) continue
      const reference = formatReference({ row: row.row, column: cell.column })
      if (cell.style !== undefined) styles.set(reference, Number(cell.style))
      if (cell.spillingFormula !== undefined) {
        sharedFormulas.set(reference, cell.spillingFormula)
      }
    }
  }

  return { styles, sharedFormulas }
}

export function sharedFormulaRefusal(reference: string, master: SpillingFormula): XlsxError {
  const what =
    master.kind === 'shared'
      ? `defines shared formula ${master.name}`
      : `holds the ${master.kind === 'array' ? 'array formula' : 'data table'} covering ${master.name}`
  return new XlsxError(
    'unwritable-value',
    `Cell ${reference} ${what}; overwriting it would break the cells that depend on it`,
    { reference },
  )
}

/**
 * An existing style is carried over so formatting survives an edit, which means
 * a Date written into a cell with no date format will show as a number.
 */
function cellElement(
  reference: string,
  value: CellInput,
  style: string | undefined,
  date1904: boolean,
  sharedStrings: ReadonlyMap<string, number> | undefined,
  prefix: string,
): string {
  checkWritable(reference, value, date1904)

  const attributes = style === undefined ? '' : ` s="${style}"`
  const c = `${prefix}c`
  const v = `${prefix}v`

  if (value === null) return `<${c} r="${reference}"${attributes}/>`

  if (typeof value === 'object' && !(value instanceof Date)) {
    // No cached result: nothing here computes one, and a stale one is worse.
    const f = `${prefix}f`
    return `<${c} r="${reference}"${attributes}><${f}>${escapeXml(value.formula)}</${f}></${c}>`
  }
  if (typeof value === 'string') {
    const shared = sharedStrings?.get(value)
    if (shared !== undefined) {
      return `<${c} r="${reference}"${attributes} t="s"><${v}>${shared}</${v}></${c}>`
    }
    const space = value === value.trim() ? '' : ' xml:space="preserve"'
    return (
      `<${c} r="${reference}"${attributes} t="inlineStr">` +
      `<${prefix}is><${prefix}t${space}>${escapeXml(value)}</${prefix}t></${prefix}is></${c}>`
    )
  }
  if (typeof value === 'boolean') {
    return `<${c} r="${reference}"${attributes} t="b"><${v}>${value ? 1 : 0}</${v}></${c}>`
  }
  if (value instanceof Date) {
    return `<${c} r="${reference}"${attributes}><${v}>${dateToSerial(value, date1904)}</${v}></${c}>`
  }
  return `<${c} r="${reference}"${attributes}><${v}>${value}</${v}></${c}>`
}

interface CellSpan {
  readonly column: number
  readonly start: number
  readonly end: number
  readonly style: string | undefined
  /** Set when other cells take their formula from this one. */
  readonly spillingFormula: SpillingFormula | undefined
}

interface RowSpan {
  readonly row: number
  readonly start: number
  /** Offset just before `</row>`, where a new cell is appended. */
  readonly contentEnd: number
  /** A self closing row holds nothing, so it is rewritten rather than spliced into. */
  readonly selfClosing: boolean
  readonly end: number
  readonly cells: CellSpan[]
}

interface DimensionSpan {
  readonly start: number
  readonly end: number
  readonly ref: string
}

interface SheetShape {
  /** Namespace prefix the document writes its elements with, `x:` or empty. */
  readonly prefix: string
  readonly dimension: DimensionSpan | undefined
  readonly rows: RowSpan[]
  /** Offset just before `</sheetData>`, where a new row is appended. */
  readonly contentEnd: number
  readonly selfClosing: boolean
  readonly dataStart: number
  readonly dataEnd: number
}

function readShape(xml: string): SheetShape {
  const rows: RowSpan[] = []
  let dimension: DimensionSpan | undefined
  let contentEnd = -1
  let dataStart = -1
  let dataEnd = -1
  let selfClosing = false

  let prefix = ''
  let lastRow = 0
  let openDimension = -1
  let currentRow: { row: number; start: number } | undefined
  let currentCells: CellSpan[] = []
  let cellStart = -1
  let cellColumn = 0
  let cellStyle: string | undefined
  let master: SpillingFormula | undefined
  let openCell = false

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.localName === 'dimension') {
        const ref = event.attributes.get('ref')
        if (ref !== undefined) {
          dimension = { start: event.start, end: event.end, ref }
          openDimension = event.selfClosing ? -1 : event.start
        }
        continue
      }
      if (event.localName === 'sheetData') {
        const colon = event.name.indexOf(':')
        prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
        dataStart = event.start
        selfClosing = event.selfClosing
        if (selfClosing) {
          dataEnd = event.end
          contentEnd = event.end
        }
        continue
      }
      if (event.localName === 'row') {
        // A row without r is the one after the row before it, which is not the
        // same as the count of rows seen once any row declares its number.
        const declared = event.attributes.get('r')
        lastRow = declared === undefined ? lastRow + 1 : Number(declared)
        currentRow = { row: lastRow, start: event.start }
        currentCells = []
        cellColumn = 0
        if (event.selfClosing) {
          rows.push({
            ...currentRow,
            contentEnd: event.end,
            end: event.end,
            selfClosing: true,
            cells: [],
          })
          currentRow = undefined
        }
        continue
      }
      if (event.localName === 'c') {
        const reference = event.attributes.get('r')
        cellColumn = reference === undefined ? cellColumn + 1 : parseReference(reference).column
        cellStyle = event.attributes.get('s')
        cellStart = event.start
        master = undefined
        openCell = !event.selfClosing
        if (event.selfClosing) {
          currentCells.push({
            column: cellColumn,
            start: event.start,
            end: event.end,
            style: cellStyle,
            spillingFormula: undefined,
          })
        }
      }
      // Only the cell carrying ref owns the range. A shared dependent names the
      // si alone, and any of them may be written self closing.
      if (event.localName === 'f') {
        const kind = event.attributes.get('t')
        const covers = event.attributes.get('ref')
        if (covers !== undefined) {
          if (kind === 'shared') {
            const si = event.attributes.get('si')
            if (si !== undefined) master = { kind: 'shared', name: si }
          } else if (kind === 'array' || kind === 'dataTable') {
            master = { kind, name: covers }
          }
        }
      }
      continue
    }

    if (event.kind !== 'close') continue

    if (event.localName === 'c' && openCell) {
      currentCells.push({
        column: cellColumn,
        start: cellStart,
        end: event.end,
        style: cellStyle,
        spillingFormula: master,
      })
      openCell = false
      continue
    }
    if (event.localName === 'row' && currentRow !== undefined) {
      rows.push({
        ...currentRow,
        contentEnd: event.start,
        end: event.end,
        selfClosing: false,
        cells: currentCells,
      })
      currentRow = undefined
      continue
    }
    if (event.localName === 'dimension' && openDimension !== -1 && dimension !== undefined) {
      // Replacing only the open tag would leave the close tag behind.
      dimension = { start: openDimension, end: event.end, ref: dimension.ref }
      openDimension = -1
      continue
    }
    if (event.localName === 'sheetData') {
      contentEnd = event.start
      dataEnd = event.end
    }
  }

  if (dataStart === -1)
    throw new XlsxError('malformed-xml', 'Sheet has no sheetData element to write into')

  return { prefix, dimension, rows, contentEnd, selfClosing, dataStart, dataEnd }
}

interface Splice {
  readonly start: number
  readonly end: number
  readonly text: string
  /** Orders insertions that land on the same offset. */
  readonly order: number
}

export function patchSheet(
  xml: string,
  edits: ReadonlyMap<string, CellInput>,
  date1904: boolean,
  sharedStrings?: ReadonlyMap<string, number>,
  styleOverrides?: ReadonlyMap<string, number>,
): string {
  if (edits.size === 0) return xml

  const shape = readShape(xml)
  const splices: Splice[] = []
  const newRows = new Map<number, string[]>()
  const filledRows = new Map<RowSpan, Array<{ column: number; cell: string }>>()

  const styleFor = (reference: string, current: string | undefined) => {
    const override = styleOverrides?.get(reference)
    return override === undefined ? current : String(override)
  }

  // Indexed rather than scanned, so writing many cells into a large sheet stays
  // linear in the number of edits instead of edits times rows.
  const rowsByNumber = new Map<number, RowSpan>()
  for (const candidate of shape.rows) rowsByNumber.set(candidate.row, candidate)

  const cellIndexes = new Map<RowSpan, Map<number, CellSpan>>()
  const cellsOf = (span: RowSpan) => {
    const known = cellIndexes.get(span)
    if (known !== undefined) return known
    const built = new Map<number, CellSpan>()
    for (const candidate of span.cells) built.set(candidate.column, candidate)
    cellIndexes.set(span, built)
    return built
  }

  for (const [given, value] of edits) {
    const address = parseReference(given)
    const { row, column } = address
    // The file never receives a reference spelled the way the caller typed it.
    const reference = formatReference(address)
    const existingRow = rowsByNumber.get(row)

    if (existingRow === undefined) {
      const pending = newRows.get(row) ?? []
      pending.push(
        cellElement(
          reference,
          value,
          styleFor(reference, undefined),
          date1904,
          sharedStrings,
          shape.prefix,
        ),
      )
      newRows.set(row, pending)
      continue
    }

    const existingCell = cellsOf(existingRow).get(column)
    if (existingCell?.spillingFormula !== undefined) {
      throw sharedFormulaRefusal(reference, existingCell.spillingFormula)
    }
    if (existingCell !== undefined) {
      splices.push({
        start: existingCell.start,
        end: existingCell.end,
        text: cellElement(
          reference,
          value,
          styleFor(reference, existingCell.style),
          date1904,
          sharedStrings,
          shape.prefix,
        ),
        order: column,
      })
      continue
    }

    const cell = cellElement(
      reference,
      value,
      styleFor(reference, undefined),
      date1904,
      sharedStrings,
      shape.prefix,
    )

    if (existingRow.selfClosing) {
      // Collected rather than spliced now: two cells added to the same empty
      // row would otherwise each rewrite it, emitting the row twice.
      const pending = filledRows.get(existingRow) ?? []
      pending.push({ column, cell })
      filledRows.set(existingRow, pending)
      continue
    }

    const next = existingRow.cells.find((candidate) => candidate.column > column)
    const at = next === undefined ? existingRow.contentEnd : next.start
    splices.push({ start: at, end: at, text: cell, order: column })
  }

  for (const [row, cells] of filledRows) {
    const ordered = [...cells]
      .sort((left, right) => left.column - right.column)
      .map((entry) => entry.cell)
      .join('')
    const openTag = xml.slice(row.start, row.end)
    splices.push({
      start: row.start,
      end: row.end,
      text: `${openTag.slice(0, -2)}>${ordered}</${shape.prefix}row>`,
      order: row.row,
    })
  }

  const buildRow = (row: number, cells: string[]) => {
    const ordered = cells
      .map((text) => ({ text, column: parseReference(cellReferenceOf(text)).column }))
      .sort((a, b) => a.column - b.column)
      .map((entry) => entry.text)
      .join('')
    return `<${shape.prefix}row r="${row}">${ordered}</${shape.prefix}row>`
  }

  if (shape.selfClosing && newRows.size > 0) {
    const body = [...newRows]
      .sort(([left], [right]) => left - right)
      .map(([row, cells]) => buildRow(row, cells))
      .join('')
    splices.push({
      start: shape.dataStart,
      end: shape.dataEnd,
      text: `<${shape.prefix}sheetData>${body}</${shape.prefix}sheetData>`,
      order: 0,
    })
  } else {
    // Existing rows come out of readShape in document order, so walking them
    // once alongside the sorted new rows places every one without rescanning
    // the sheet per row.
    let at = 0
    let next = shape.rows[at]
    for (const [row, cells] of [...newRows].sort(([left], [right]) => left - right)) {
      while (next !== undefined && next.row <= row) {
        at++
        next = shape.rows[at]
      }
      const offset = next === undefined ? shape.contentEnd : next.start
      splices.push({ start: offset, end: offset, text: buildRow(row, cells), order: row })
    }
  }

  const widened = widenDimension(shape.dimension, [...edits.keys()])
  if (widened !== undefined && shape.dimension !== undefined) {
    splices.push({
      start: shape.dimension.start,
      end: shape.dimension.end,
      text: `<${shape.prefix}dimension ref="${widened}"/>`,
      order: -1,
    })
  }

  return applySplices(xml, splices)
}

/**
 * Excel recalculates the used range, but stricter readers trust what the file
 * declares, so a cell written outside it would be ignored.
 */
function widenDimension(
  dimension: DimensionSpan | undefined,
  references: readonly string[],
): string | undefined {
  if (dimension === undefined) return undefined

  const bounds = dimension.ref.split(':')
  const from = bounds[0] ?? ''
  if (from === '') return undefined

  const topLeft = parseReference(from)
  const bottomRight = parseReference(bounds[1] ?? from)

  let { row: lastRow, column: lastColumn } = bottomRight
  let { row: firstRow, column: firstColumn } = topLeft

  for (const reference of references) {
    const { row, column } = parseReference(reference)
    firstRow = Math.min(firstRow, row)
    firstColumn = Math.min(firstColumn, column)
    lastRow = Math.max(lastRow, row)
    lastColumn = Math.max(lastColumn, column)
  }

  const grown =
    firstRow !== topLeft.row ||
    firstColumn !== topLeft.column ||
    lastRow !== bottomRight.row ||
    lastColumn !== bottomRight.column
  if (!grown) return undefined

  const start = formatReference({ row: firstRow, column: firstColumn })
  const end = formatReference({ row: lastRow, column: lastColumn })
  return `${start}:${end}`
}

/** Only safe on elements built above, where the reference is the first attribute. */
function cellReferenceOf(element: string): string {
  const start = element.indexOf('"') + 1
  return element.slice(start, element.indexOf('"', start))
}

function applySplices(xml: string, splices: Splice[]): string {
  const ordered = [...splices].sort((a, b) => a.start - b.start || a.order - b.order)

  const pieces: string[] = []
  let cursor = 0
  for (const splice of ordered) {
    pieces.push(xml.slice(cursor, splice.start))
    pieces.push(splice.text)
    cursor = splice.end
  }
  pieces.push(xml.slice(cursor))

  return pieces.join('')
}
