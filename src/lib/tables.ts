import type { Container } from './container.js'
import { formatReference, parseReference } from './reference.js'
import { readRelationships, resolveTarget } from './relationships.js'
import { readXml } from './xml.js'

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
 * Excel grows a table to cover a cell written directly below it. Returns the
 * rewritten parts for a sheet's tables that an edit extended; a table no edit
 * reaches, or one this slice does not handle, is left exactly as it was.
 */
export function extendTables(
  sheetXml: string,
  sheetPath: string,
  container: Container,
  written: Iterable<string>,
): TableExtension[] {
  const paths = tablePartPaths(sheetXml, sheetPath, container)
  if (paths.length === 0) return []

  const cells = new Set(written)
  const tables = paths
    .map((path) => decodeTable(container, path))
    .filter((table): table is DecodedTable => table !== undefined)
  const ranges = tables.map((table) => table.range)

  const extensions: TableExtension[] = []
  for (const table of tables) {
    const grown = growDown(table, cells, ranges)
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
}

function tablePartPaths(sheetXml: string, sheetPath: string, container: Container): string[] {
  const ids: string[] = []
  for (const event of readXml(sheetXml)) {
    if (event.kind !== 'open' || event.localName !== 'tablePart') continue
    const id = relationshipId(event.attributes)
    if (id !== undefined) ids.push(id)
  }
  if (ids.length === 0) return []

  const relsPath = sheetPath.replace(/([^/]+)$/, '_rels/$1.rels')
  const relsBytes = container.parts.get(relsPath)
  if (relsBytes === undefined) return []
  const relationships = readRelationships(decode(relsBytes), relsPath)

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

  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'table') continue
    const ref = event.attributes.get('ref')
    const range = ref === undefined ? undefined : parseRange(ref)
    if (range === undefined) return undefined
    // totalsRowShown defaults to true and only says the row was once visible;
    // the count is what says a totals row exists and occupies the row below.
    const totals = event.attributes.get('totalsRowCount')
    const hasTotalsRow = totals !== undefined && totals !== '0'
    return { path, xml, range, hasTotalsRow }
  }
  return undefined
}

function parseRange(ref: string): Range | undefined {
  const colon = ref.indexOf(':')
  if (colon === -1) return undefined
  const start = parseReference(ref.slice(0, colon))
  const end = parseReference(ref.slice(colon + 1))
  return {
    minRow: Math.min(start.row, end.row),
    maxRow: Math.max(start.row, end.row),
    minColumn: Math.min(start.column, end.column),
    maxColumn: Math.max(start.column, end.column),
  }
}

function growDown(
  table: DecodedTable,
  written: Set<string>,
  all: readonly Range[],
): string | undefined {
  if (table.hasTotalsRow) return undefined
  const { range } = table

  const columnWritten = (row: number) => {
    for (let column = range.minColumn; column <= range.maxColumn; column++) {
      if (written.has(formatReference({ row, column }))) return true
    }
    return false
  }
  // A grown row must not land inside another table that shares any column.
  const others = all.filter((candidate) => candidate !== range)
  const collides = (row: number) =>
    others.some(
      (other) =>
        row >= other.minRow &&
        row <= other.maxRow &&
        range.minColumn <= other.maxColumn &&
        range.maxColumn >= other.minColumn,
    )

  let maxRow = range.maxRow
  while (columnWritten(maxRow + 1) && !collides(maxRow + 1)) maxRow++
  if (maxRow === range.maxRow) return undefined

  const oldRef = `${formatReference({ row: range.minRow, column: range.minColumn })}:${formatReference({ row: range.maxRow, column: range.maxColumn })}`
  const newRef = `${formatReference({ row: range.minRow, column: range.minColumn })}:${formatReference({ row: maxRow, column: range.maxColumn })}`
  return rewriteRefs(table.xml, oldRef, newRef)
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
