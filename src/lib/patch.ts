import { dateToSerial } from './date.js'
import { XlsxError } from './errors.js'
import { parseReference } from './reference.js'
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
): string {
  const attributes = style === undefined ? '' : ` s="${style}"`

  if (value === null) return `<c r="${reference}"${attributes}/>`
  if (typeof value === 'string') {
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
  readonly cells: CellSpan[]
}

interface SheetShape {
  readonly rows: RowSpan[]
  /** Offset just before `</sheetData>`, where a new row is appended. */
  readonly contentEnd: number
  readonly selfClosing: boolean
  readonly dataStart: number
  readonly dataEnd: number
}

function readShape(xml: string): SheetShape {
  const rows: RowSpan[] = []
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
        if (event.selfClosing) {
          rows.push({ ...currentRow, contentEnd: event.end, cells: [] })
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
      rows.push({ ...currentRow, contentEnd: event.start, cells: currentCells })
      currentRow = undefined
      continue
    }
    if (event.name === 'sheetData') {
      contentEnd = event.start
      dataEnd = event.end
    }
  }

  if (dataStart === -1) throw new XlsxError('Sheet has no sheetData element to write into')

  return { rows, contentEnd, selfClosing, dataStart, dataEnd }
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
): string {
  if (edits.size === 0) return xml

  const shape = readShape(xml)
  const splices: Splice[] = []
  const newRows = new Map<number, string[]>()

  for (const [reference, value] of edits) {
    const { row, column } = parseReference(reference)
    const existingRow = shape.rows.find((candidate) => candidate.row === row)

    if (existingRow === undefined) {
      const pending = newRows.get(row) ?? []
      pending.push(cellElement(reference, value, undefined, date1904))
      newRows.set(row, pending)
      continue
    }

    const existingCell = existingRow.cells.find((candidate) => candidate.column === column)
    if (existingCell !== undefined) {
      splices.push({
        start: existingCell.start,
        end: existingCell.end,
        text: cellElement(reference, value, existingCell.style, date1904),
        order: column,
      })
      continue
    }

    const next = existingRow.cells.find((candidate) => candidate.column > column)
    const at = next === undefined ? existingRow.contentEnd : next.start
    splices.push({
      start: at,
      end: at,
      text: cellElement(reference, value, undefined, date1904),
      order: column,
    })
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

  return applySplices(xml, splices)
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
