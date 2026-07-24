import { type Container, decodeXmlPart } from './container.js'
import { formatReference, parseReference } from './reference.js'
import { readRelationships, resolveTarget } from './relationships.js'
import { readXml, readXmlBytes, withAttribute } from './xml.js'

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

interface DecodedTable {
  readonly path: string
  readonly xml: string
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
  if (bytes === undefined) return undefined
  let xml: string
  try {
    xml = decode(bytes)
  } catch {
    return undefined
  }

  let range: Range | undefined
  let hasTotalsRow = false
  const columnNames: string[] = []
  const columnIds: number[] = []

  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    if (event.localName === 'table') {
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
  return { path, xml, range, hasTotalsRow, columnNames, columnIds }
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
