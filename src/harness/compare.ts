import { readParts, bytesEqual } from './zip.js'
import { scanFeatures } from './features.js'
import type { Adapter, CellValue, RoundTripResult, FeatureDelta, SheetValues } from './types.js'

/**
 * Parts that legitimately differ on every write and are not evidence of
 * data loss: document timestamps, generator metadata, and the calculation
 * chain (which Excel rebuilds anyway). Directory entries are an artifact of
 * how the zip was produced and carry no content.
 */
const VOLATILE = [/^docProps\/core\.xml$/, /^docProps\/app\.xml$/, /^xl\/calcChain\.xml$/]

const isVolatile = (path: string) =>
  path.endsWith('/') || VOLATILE.some((pattern) => pattern.test(path))

/**
 * Writing a cell has to rewrite the sheet it landed in, and may have to add a
 * shared string, a cell format, or a content type. Everything else changing is
 * the library reaching further into the document than the edit called for.
 */
const EDIT_MAY_CHANGE = [
  /^\[Content_Types\]\.xml$/,
  /^xl\/workbook\.xml$/,
  /^xl\/styles\.xml$/,
  /^xl\/sharedStrings\.xml$/,
  // The calculation chain is dropped on an edit, so the relationship naming it
  // has to go with it. Leaving it behind is an invalid package.
  /^xl\/_rels\/workbook\.xml\.rels$/,
  // A cell written just past a table grows it, so its part changes. Loss would
  // still show as a dropped table part or a changed cell, which are caught.
  /^xl\/tables\/[^/]+\.xml$/,
  // Adding a hyperlink gives it a relationship, so the sheet's rels change. A
  // relationship dropped alongside would surface as its target part going
  // missing, which is caught.
  /^xl\/worksheets\/_rels\/[^/]+\.xml\.rels$/,
]

const isWorksheet = (path: string) => /^xl\/worksheets\/[^/]+\.xml$/.test(path)

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export async function measureRoundTrip(
  adapter: Adapter,
  file: string,
  bytes: Uint8Array,
  edited = false,
): Promise<RoundTripResult> {
  const base: RoundTripResult = {
    file,
    ok: false,
    partsLost: [],
    partsAdded: [],
    partsChanged: [],
    featureLoss: [],
    cellsLost: 0,
    cellsChanged: 0,
    cellsTotal: 0,
    bytesIn: bytes.length,
    bytesOut: 0,
  }

  const produce = edited ? adapter.edit : adapter.roundTrip
  if (produce === undefined) return { ...base, error: 'adapter cannot edit' }

  let out: Uint8Array
  try {
    out = await produce.call(adapter, bytes)
  } catch (error) {
    return { ...base, error: `${edited ? 'edit' : 'roundTrip'}: ${describeError(error)}` }
  }
  base.bytesOut = out.length

  const before = readParts(bytes)
  let after: Map<string, Uint8Array>
  try {
    after = readParts(out)
  } catch (error) {
    return { ...base, error: `output is not a readable zip: ${describeError(error)}` }
  }

  const changedSheets: string[] = []
  for (const [path, content] of before) {
    if (isVolatile(path)) continue
    const other = after.get(path)
    if (!other) {
      base.partsLost.push(path)
      continue
    }
    if (bytesEqual(content, other)) continue
    if (edited) {
      if (EDIT_MAY_CHANGE.some((pattern) => pattern.test(path))) continue
      if (isWorksheet(path)) {
        changedSheets.push(path)
        continue
      }
    }
    base.partsChanged.push(path)
  }
  // One edit lands in one sheet. More than that is the write path spilling.
  if (changedSheets.length > 1) base.partsChanged.push(...changedSheets)
  for (const path of after.keys()) {
    if (!before.has(path) && !isVolatile(path)) base.partsAdded.push(path)
  }

  base.featureLoss = diffFeatures(scanFeatures(before), scanFeatures(after))

  try {
    const cells = await compareValues(adapter, bytes, out)
    Object.assign(base, cells)
  } catch (error) {
    return { ...base, error: `value comparison: ${describeError(error)}` }
  }

  // ok means no data was lost. A part whose bytes merely differ is reported
  // separately: rewriting XML is legitimate for a library that reserialises,
  // and only this one claims to leave untouched parts alone.
  base.ok =
    base.partsLost.length === 0 &&
    base.featureLoss.length === 0 &&
    base.cellsLost === 0 &&
    base.cellsChanged === 0

  return base
}

function diffFeatures(before: Map<string, number>, after: Map<string, number>): FeatureDelta[] {
  const deltas: FeatureDelta[] = []
  for (const [feature, count] of before) {
    const now = after.get(feature) ?? 0
    if (now < count) deltas.push({ feature, before: count, after: now })
  }
  return deltas.sort((a, b) => b.before - b.after - (a.before - a.after))
}

async function compareValues(
  adapter: Adapter,
  original: Uint8Array,
  roundTripped: Uint8Array,
): Promise<Pick<RoundTripResult, 'cellsLost' | 'cellsChanged' | 'cellsTotal'>> {
  const before = byQualifiedAddress(await adapter.values(original))
  const after = byQualifiedAddress(await adapter.values(roundTripped))

  let cellsLost = 0
  let cellsChanged = 0
  for (const [key, value] of before) {
    const other = after.get(key)
    if (other === undefined) cellsLost++
    else if (!sameValue(value, other)) cellsChanged++
  }

  return { cellsLost, cellsChanged, cellsTotal: before.size }
}

function byQualifiedAddress(sheets: SheetValues[]): Map<string, CellValue> {
  const flat = new Map<string, CellValue>()
  for (const sheet of sheets) {
    for (const [address, value] of sheet.cells) {
      flat.set(`${sheet.name}!${address}`, value)
    }
  }
  return flat
}

function sameValue(a: CellValue, b: CellValue): boolean {
  if (a.type !== b.type) return false
  if (a.formula !== b.formula) return false
  if (a.style !== b.style) return false
  if (typeof a.value === 'number' && typeof b.value === 'number') {
    // Spreadsheet numbers round-trip through decimal text; tolerate the last bit.
    return Math.abs(a.value - b.value) < 1e-9
  }
  return a.value === b.value
}
