import { dateToSerial } from './date.js'
import { XlsxError } from './errors.js'
import { formatReference, parseReference } from './reference.js'
import { readXml } from './xml.js'

export type CellInput = number | string | boolean | Date | null

const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

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
): string {
  const attributes = style === undefined ? '' : ` s="${style}"`

  if (value === null) return `<c r="${reference}"${attributes}/>`
  if (typeof value === 'string') {
    const shared = sharedStrings?.get(value)
    if (shared !== undefined) {
      return `<c r="${reference}"${attributes} t="s"><v>${shared}</v></c>`
    }
    return `<c r="${reference}"${attributes} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`
  }
  if (typeof value === 'boolean') {
    return `<c r="${reference}"${attributes} t="b"><v>${value ? 1 : 0}</v></c>`
  }
  if (value instanceof Date) {
    return `<c r="${reference}"${attributes}><v>${dateToSerial(value, date1904)}</v></c>`
  }
  if (!Number.isFinite(value)) {
    throw new XlsxError(`Cell ${reference} cannot hold ${value}`)
  }
  return `<c r="${reference}"${attributes}><v>${value}</v></c>`
}

interface CellSpan {
  readonly column: number
  readonly start: number
  readonly end: number
  readonly style: string | undefined
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

  let currentRow: { row: number; start: number } | undefined
  let currentCells: CellSpan[] = []
  let cellStart = -1
  let cellColumn = 0
  let cellStyle: string | undefined
  let openCell = false

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.name === 'dimension') {
        const ref = event.attributes.get('ref')
        if (ref !== undefined) dimension = { start: event.start, end: event.end, ref }
        continue
      }
      if (event.name === 'sheetData') {
        dataStart = event.start
        selfClosing = event.selfClosing
        if (selfClosing) {
          dataEnd = event.end
          contentEnd = event.end
        }
        continue
      }
      if (event.name === 'row') {
        const declared = event.attributes.get('r')
        currentRow = {
          row: declared === undefined ? rows.length + 1 : Number(declared),
          start: event.start,
        }
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
      if (event.name === 'c') {
        const reference = event.attributes.get('r')
        cellColumn = reference === undefined ? cellColumn + 1 : parseReference(reference).column
        cellStyle = event.attributes.get('s')
        cellStart = event.start
        openCell = !event.selfClosing
        if (event.selfClosing) {
          currentCells.push({
            column: cellColumn,
            start: event.start,
            end: event.end,
            style: cellStyle,
          })
        }
      }
      continue
    }

    if (event.kind !== 'close') continue

    if (event.name === 'c' && openCell) {
      currentCells.push({ column: cellColumn, start: cellStart, end: event.end, style: cellStyle })
      openCell = false
      continue
    }
    if (event.name === 'row' && currentRow !== undefined) {
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
    if (event.name === 'sheetData') {
      contentEnd = event.start
      dataEnd = event.end
    }
  }

  if (dataStart === -1) throw new XlsxError('Sheet has no sheetData element to write into')

  return { dimension, rows, contentEnd, selfClosing, dataStart, dataEnd }
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

  const styleFor = (reference: string, current: string | undefined) => {
    const override = styleOverrides?.get(reference)
    return override === undefined ? current : String(override)
  }

  for (const [given, value] of edits) {
    const address = parseReference(given)
    const { row, column } = address
    // The file never receives a reference spelled the way the caller typed it.
    const reference = formatReference(address)
    const existingRow = shape.rows.find((candidate) => candidate.row === row)

    if (existingRow === undefined) {
      const pending = newRows.get(row) ?? []
      pending.push(
        cellElement(reference, value, styleFor(reference, undefined), date1904, sharedStrings),
      )
      newRows.set(row, pending)
      continue
    }

    const existingCell = existingRow.cells.find((candidate) => candidate.column === column)
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
    )

    if (existingRow.selfClosing) {
      const openTag = xml.slice(existingRow.start, existingRow.end)
      splices.push({
        start: existingRow.start,
        end: existingRow.end,
        text: `${openTag.slice(0, -2)}>${cell}</row>`,
        order: column,
      })
      continue
    }

    const next = existingRow.cells.find((candidate) => candidate.column > column)
    const at = next === undefined ? existingRow.contentEnd : next.start
    splices.push({ start: at, end: at, text: cell, order: column })
  }

  const buildRow = (row: number, cells: string[]) => {
    const ordered = cells
      .map((text) => ({ text, column: parseReference(cellReferenceOf(text)).column }))
      .sort((a, b) => a.column - b.column)
      .map((entry) => entry.text)
      .join('')
    return `<row r="${row}">${ordered}</row>`
  }

  if (shape.selfClosing && newRows.size > 0) {
    const body = [...newRows]
      .sort(([left], [right]) => left - right)
      .map(([row, cells]) => buildRow(row, cells))
      .join('')
    splices.push({
      start: shape.dataStart,
      end: shape.dataEnd,
      text: `<sheetData>${body}</sheetData>`,
      order: 0,
    })
  } else {
    for (const [row, cells] of newRows) {
      const next = shape.rows.find((candidate) => candidate.row > row)
      const at = next === undefined ? shape.contentEnd : next.start
      splices.push({ start: at, end: at, text: buildRow(row, cells), order: row })
    }
  }

  const widened = widenDimension(shape.dimension, [...edits.keys()])
  if (widened !== undefined && shape.dimension !== undefined) {
    splices.push({
      start: shape.dimension.start,
      end: shape.dimension.end,
      text: `<dimension ref="${widened}"/>`,
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
