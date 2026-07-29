import { escapeSheetName } from './add-sheet.js'
import { type Container, decodeXmlPart } from './container.js'
import { formatReference, parseReference } from './reference.js'
import { readRelationships, resolveTarget } from './relationships.js'
import type { ShiftSpec } from './shift.js'
import { readXml, readXmlBytes, withAttribute } from './xml.js'

const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** A table to author: its name, the range it covers, its column headers and the
 * built-in style to draw it with. `name` is Excel's table name and display name. */
export interface TableSpec {
  readonly name: string
  readonly ref: string
  readonly columns: readonly string[]
  readonly style: string
}

/** Builds a table part. Columns are numbered from one; the auto-filter starts out
 * equal to the table's own range, the way Excel writes a fresh table. */
export function buildTablePart(id: number, spec: TableSpec): string {
  const columns = spec.columns
    .map((name, index) => `<tableColumn id="${index + 1}" name="${escapeSheetName(name)}"/>`)
    .join('')
  const name = escapeSheetName(spec.name)
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<table xmlns="${SPREADSHEET_NS}" id="${id}" name="${name}" displayName="${name}" ` +
    `ref="${spec.ref}" totalsRowShown="0"><autoFilter ref="${spec.ref}"/>` +
    `<tableColumns count="${spec.columns.length}">${columns}</tableColumns>` +
    `<tableStyleInfo name="${escapeSheetName(spec.style)}" showFirstColumn="0" ` +
    'showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
  )
}

/**
 * Wires a table into a sheet through a `<tableParts>` element, joining one the
 * sheet already has. Placed before a worksheet-level `extLst` or the closing tag,
 * where the schema puts it. Declares `xmlns:r` on the worksheet when the sheet —
 * a blank one this library created, say — does not already carry it.
 */
export function withTableParts(sheetXml: string, relationshipId: string): string {
  let prefix = ''
  let worksheet: { start: number; end: number } | undefined
  let container: { openStart: number; openEnd: number; insertAt: number; count: number } | undefined
  let selfClosing = false
  let worksheetClose = -1
  let extLst = -1
  for (const event of readXml(sheetXml)) {
    if (event.kind === 'open' && event.localName === 'worksheet') {
      const colon = event.name.indexOf(':')
      prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
      worksheet = { start: event.start, end: event.end }
    } else if (event.kind === 'open' && event.localName === 'tableParts') {
      selfClosing = event.selfClosing
      container = {
        openStart: event.start,
        openEnd: event.end,
        insertAt: event.end,
        count: Number(event.attributes.get('count')) || 0,
      }
    } else if (
      event.kind === 'close' &&
      event.localName === 'tableParts' &&
      container !== undefined
    ) {
      container = { ...container, insertAt: event.start }
    } else if (event.kind === 'open' && event.localName === 'extLst' && extLst === -1) {
      extLst = event.start
    } else if (event.kind === 'close' && event.localName === 'worksheet') {
      worksheetClose = event.start
    }
  }

  const child = `<${prefix}tablePart r:id="${relationshipId}"/>`
  const splices: { start: number; end: number; text: string }[] = []
  if (
    worksheet !== undefined &&
    !sheetXml.slice(worksheet.start, worksheet.end).includes('xmlns:r=')
  ) {
    const tag = sheetXml.slice(worksheet.start, worksheet.end)
    splices.push({
      start: worksheet.start,
      end: worksheet.end,
      text: tag.replace(/^<([^\s/>]+)/, `<$1 xmlns:r="${RELATIONSHIPS_NS}"`),
    })
  }

  if (container === undefined) {
    const anchor = extLst !== -1 ? extLst : worksheetClose !== -1 ? worksheetClose : sheetXml.length
    splices.push({
      start: anchor,
      end: anchor,
      text: `<${prefix}tableParts count="1">${child}</${prefix}tableParts>`,
    })
  } else {
    const counted = withAttribute(
      sheetXml.slice(container.openStart, container.openEnd),
      'count',
      container.count + 1,
    )
    if (selfClosing) {
      splices.push({
        start: container.openStart,
        end: container.openEnd,
        text: `${counted.slice(0, -2)}>${child}</${prefix}tableParts>`,
      })
    } else {
      splices.push({ start: container.openStart, end: container.openEnd, text: counted })
      splices.push({ start: container.insertAt, end: container.insertAt, text: child })
    }
  }

  let xml = sheetXml
  for (const splice of splices.sort((a, b) => b.start - a.start)) {
    xml = xml.slice(0, splice.start) + splice.text + xml.slice(splice.end)
  }
  return xml
}

/** The tables a sheet carries, each with its range and column names. */
export function readTables(
  sheetBytes: Uint8Array,
  sheetPath: string,
  container: Container,
): { name: string; range: string; columns: string[] }[] {
  const tables: { name: string; range: string; columns: string[] }[] = []
  for (const path of tablePartPaths(sheetBytes, sheetPath, container)) {
    const table = decodeTable(container, path)
    if (table !== undefined && table.name !== undefined) {
      tables.push({ name: table.name, range: refOf(table.range), columns: table.columnNames })
    }
  }
  return tables
}

/** A rewritten table part, ready to replace the one that was read. */
export interface TableExtension {
  readonly path: string
  readonly xml: string
}

interface Range {
  readonly minRow: number
  readonly maxRow: number
  readonly minColumn: number
  readonly maxColumn: number
}

/**
 * Excel grows a table to cover a cell written directly below or to the right of
 * it. Returns the rewritten parts for a sheet's tables that an edit extended; a
 * table no edit reaches, or one with a totals row, is left exactly as it was.
 */
export function extendTables(
  sheetBytes: Uint8Array,
  sheetPath: string,
  container: Container,
  written: Iterable<string>,
): TableExtension[] {
  const paths = tablePartPaths(sheetBytes, sheetPath, container)
  if (paths.length === 0) return []

  const cells = new Set(written)
  const tables = paths
    .map((path) => decodeTable(container, path))
    .filter((table): table is DecodedTable => table !== undefined)
  const ranges = tables.map((table) => table.range)

  const extensions: TableExtension[] = []
  for (const table of tables) {
    const grown = grow(table, cells, ranges)
    if (grown !== undefined) extensions.push({ path: table.path, xml: grown })
  }
  return extensions
}

// A line at or past the insert/delete point moves; a deletion clamps a line that
// falls inside the removed band to the surviving edge — the low end up to the row
// after it, the high end down to the row before — so the range only shrinks.
const shiftLow = (value: number, spec: ShiftSpec): number => {
  if (value < spec.at) return value
  if (spec.delta > 0) return value + spec.delta
  return value < spec.at - spec.delta ? spec.at : value + spec.delta
}
const shiftHigh = (value: number, spec: ShiftSpec): number => {
  if (value < spec.at) return value
  if (spec.delta > 0) return value + spec.delta
  return value < spec.at - spec.delta ? spec.at - 1 : value + spec.delta
}

/**
 * The rewritten table parts for a sheet whose rows an insert or delete moved.
 * Only the row axis adjusts a table so far; a column edit does not reach here.
 * `currentPart` gives the latest bytes for a table already rewritten this
 * session — a grow before an insert — so a chain of edits composes.
 */
export function shiftTables(
  currentPart: (path: string) => Uint8Array | undefined,
  sheetBytes: Uint8Array,
  sheetPath: string,
  container: Container,
  spec: ShiftSpec,
): TableExtension[] {
  if (spec.axis !== 'row') return []
  const extensions: TableExtension[] = []
  for (const path of tablePartPaths(sheetBytes, sheetPath, container)) {
    const bytes = currentPart(path) ?? container.parts.get(path)
    const table = bytes === undefined ? undefined : decodeTableBytes(bytes, path)
    if (table === undefined) continue
    const shifted = {
      ...table.range,
      minRow: shiftLow(table.range.minRow, spec),
      maxRow: shiftHigh(table.range.maxRow, spec),
    }
    if (shifted.minRow === table.range.minRow && shifted.maxRow === table.range.maxRow) continue
    extensions.push({ path, xml: rewriteRefs(table.xml, refOf(table.range), refOf(shifted)) })
  }
  return extensions
}

/**
 * The name of a table a row deletion would destroy by taking its header row —
 * its band covers the row the table starts on, so there is nothing to shrink to.
 * Only reached from `deleteRows`, so the spec is always a row removal.
 */
export function tableRowDamage(
  sheetBytes: Uint8Array,
  sheetPath: string,
  container: Container,
  spec: ShiftSpec,
): string | undefined {
  for (const path of tablePartPaths(sheetBytes, sheetPath, container)) {
    const table = decodeTable(container, path)
    if (table === undefined) continue
    const header = table.range.minRow
    if (spec.at <= header && header < spec.at - spec.delta) {
      return table.name === undefined ? 'a table' : `table ${table.name}`
    }
  }
  return undefined
}

interface DecodedTable {
  readonly path: string
  readonly xml: string
  readonly name: string | undefined
  readonly range: Range
  /** A totals row sits below the data, so the row below the table is not free. */
  readonly hasTotalsRow: boolean
  /** Existing column names and ids, so an added column can avoid both. */
  readonly columnNames: string[]
  readonly columnIds: number[]
}

function tablePartPaths(sheetBytes: Uint8Array, sheetPath: string, container: Container): string[] {
  const ids: string[] = []
  for (const event of readXmlBytes(sheetBytes)) {
    if (event.kind !== 'open' || event.localName !== 'tablePart') continue
    const id = relationshipId(event.attributes)
    if (id !== undefined) ids.push(id)
  }
  if (ids.length === 0) return []

  const relsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const relsBytes = container.parts.get(relsPath)
  if (relsBytes === undefined) return []
  const relationships = readRelationships(decodeXmlPart(relsBytes, relsPath), relsPath)

  const paths: string[] = []
  for (const id of ids) {
    const target = relationships.get(id)?.target
    if (target !== undefined) paths.push(resolveTarget(sheetPath, target))
  }
  return paths
}

const decode = (bytes: Uint8Array) => new TextDecoder('utf-8', { fatal: true }).decode(bytes)

function decodeTable(container: Container, path: string): DecodedTable | undefined {
  const bytes = container.parts.get(path)
  return bytes === undefined ? undefined : decodeTableBytes(bytes, path)
}

function decodeTableBytes(bytes: Uint8Array, path: string): DecodedTable | undefined {
  let xml: string
  try {
    xml = decode(bytes)
  } catch {
    return undefined
  }

  let range: Range | undefined
  let name: string | undefined
  let hasTotalsRow = false
  const columnNames: string[] = []
  const columnIds: number[] = []

  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    if (event.localName === 'table') {
      name = event.attributes.get('displayName') ?? event.attributes.get('name')
      const ref = event.attributes.get('ref')
      range = ref === undefined ? undefined : parseRange(ref)
      if (range === undefined) return undefined
      // totalsRowShown defaults to true and only says the row was once visible;
      // the count is what says a totals row exists and occupies the row below.
      const totals = event.attributes.get('totalsRowCount')
      hasTotalsRow = totals !== undefined && totals !== '0'
    }
    if (event.localName === 'tableColumn') {
      const name = event.attributes.get('name')
      if (name !== undefined) columnNames.push(name)
      const id = Number(event.attributes.get('id'))
      if (Number.isInteger(id)) columnIds.push(id)
    }
  }

  if (range === undefined) return undefined
  return { path, xml, name, range, hasTotalsRow, columnNames, columnIds }
}

function parseRange(ref: string): Range | undefined {
  const colon = ref.indexOf(':')
  if (colon === -1) return undefined
  // A malformed half is a fault in the file, not the caller, so the table is
  // left as it is rather than taking the whole save down with a bad-reference.
  try {
    const start = parseReference(ref.slice(0, colon))
    const end = parseReference(ref.slice(colon + 1))
    return {
      minRow: Math.min(start.row, end.row),
      maxRow: Math.max(start.row, end.row),
      minColumn: Math.min(start.column, end.column),
      maxColumn: Math.max(start.column, end.column),
    }
  } catch {
    return undefined
  }
}

function grow(
  table: DecodedTable,
  written: Set<string>,
  all: readonly Range[],
): string | undefined {
  // A totals row, or an added column beside one, needs a totals cell we do not
  // write, so a table that has one is left exactly as it was.
  if (table.hasTotalsRow) return undefined
  const { range } = table
  const others = all.filter((candidate) => candidate !== range)

  const rowWritten = (row: number) =>
    spanWritten(written, row, row, range.minColumn, range.maxColumn)
  const columnHasWrite = (column: number) =>
    spanWritten(written, range.minRow, range.maxRow, column, column)
  const overlaps = (r: Range) =>
    others.some(
      (other) =>
        r.minRow <= other.maxRow &&
        r.maxRow >= other.minRow &&
        r.minColumn <= other.maxColumn &&
        r.maxColumn >= other.minColumn,
    )
  const downTo = (row: number) => ({ ...range, maxRow: row })
  const rightTo = (column: number) => ({ ...range, maxColumn: column })

  let maxRow = range.maxRow
  while (rowWritten(maxRow + 1) && !overlaps(downTo(maxRow + 1))) maxRow++
  let maxColumn = range.maxColumn
  while (columnHasWrite(maxColumn + 1) && !overlaps(rightTo(maxColumn + 1))) maxColumn++

  if (maxRow === range.maxRow && maxColumn === range.maxColumn) return undefined

  const oldRef = refOf(range)
  const newRef = refOf({ ...range, maxRow, maxColumn })
  let xml = rewriteRefs(table.xml, oldRef, newRef)
  if (maxColumn > range.maxColumn) {
    xml = addColumns(xml, maxColumn - range.maxColumn, table.columnIds, table.columnNames)
  }
  return xml
}

const refOf = (range: Range) =>
  `${formatReference({ row: range.minRow, column: range.minColumn })}:${formatReference({ row: range.maxRow, column: range.maxColumn })}`

function spanWritten(
  written: Set<string>,
  minRow: number,
  maxRow: number,
  minColumn: number,
  maxColumn: number,
): boolean {
  for (let row = minRow; row <= maxRow; row++) {
    for (let column = minColumn; column <= maxColumn; column++) {
      if (written.has(formatReference({ row, column }))) return true
    }
  }
  return false
}

/**
 * Excel names a fresh column `ColumnN`, so we take the first free of both,
 * resuming from `from` rather than rescanning from 1: a name once taken stays
 * taken through the batch, so the smallest free index never moves backwards.
 */
function freshColumnName(taken: Set<string>, from: number): { name: string; next: number } {
  let n = from
  while (taken.has(`column${n}`)) n++
  return { name: `Column${n}`, next: n + 1 }
}

/**
 * Appends `count` empty columns to the table's tableColumns, each with an id and
 * a name no existing column uses, and raises the declared count to match.
 */
function addColumns(xml: string, count: number, ids: number[], names: string[]): string {
  const taken = new Set(names.map((name) => name.toLowerCase()))
  // A fold, not Math.max(...ids): spreading a file-sized id array as call
  // arguments overflows the stack on a table with hundreds of thousands of columns.
  let nextId = ids.reduce((max, id) => Math.max(max, id), 0) + 1

  const added: string[] = []
  let cursor = 1
  for (let index = 0; index < count; index++) {
    const { name, next } = freshColumnName(taken, cursor)
    cursor = next
    taken.add(name.toLowerCase())
    added.push(`<tableColumn id="${nextId++}" name="${name}"/>`)
  }
  const elements = added.join('')

  let prefix = ''
  let insertAt = -1
  let countStart = -1
  let countEnd = -1
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'tableColumns') {
      const colon = event.name.indexOf(':')
      prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
      countStart = event.start
      countEnd = event.end
    }
    if (event.kind === 'close' && event.localName === 'tableColumns') insertAt = event.start
  }
  if (insertAt === -1 || countStart === -1) return xml

  const openTag = withAttribute(xml.slice(countStart, countEnd), 'count', names.length + count)
  const prefixed = elements.replace(/<tableColumn /g, `<${prefix}tableColumn `)
  return (
    xml.slice(0, countStart) +
    openTag +
    xml.slice(countEnd, insertAt) +
    prefixed +
    xml.slice(insertAt)
  )
}

/**
 * Replaces the table's own ref and its autoFilter's, which start out equal. The
 * autoFilter is only touched when it matches, so a filter over a sub-range is
 * left alone rather than silently widened.
 */
function rewriteRefs(xml: string, oldRef: string, newRef: string): string {
  let out = ''
  let position = 0
  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    if (event.localName !== 'table' && event.localName !== 'autoFilter') continue

    const tag = xml.slice(event.start, event.end)
    const replaced = tag.replace(
      /(\sref\s*=\s*)("[^"]*"|'[^']*')/,
      (whole, head: string, quoted: string) => {
        const value = quoted.slice(1, -1)
        if (value !== oldRef) return whole
        const quote = quoted.charAt(0)
        return `${head}${quote}${newRef}${quote}`
      },
    )
    out += xml.slice(position, event.start) + replaced
    position = event.end
  }
  return out + xml.slice(position)
}

/** The namespace prefix is the file's choice, so `r:` cannot be assumed. */
function relationshipId(attributes: ReadonlyMap<string, string>): string | undefined {
  const direct = attributes.get('r:id')
  if (direct !== undefined) return direct
  for (const [name, value] of attributes) {
    if (name === 'id' || name.endsWith(':id')) return value
  }
  return undefined
}
