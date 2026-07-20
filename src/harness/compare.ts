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

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

export async function measureRoundTrip(
  adapter: Adapter,
  file: string,
  bytes: Uint8Array,
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

  let out: Uint8Array
  try {
    out = await adapter.roundTrip(bytes)
  } catch (error) {
    return { ...base, error: `roundTrip: ${describeError(error)}` }
  }
  base.bytesOut = out.length

  const before = readParts(bytes)
  let after: Map<string, Uint8Array>
  try {
    after = readParts(out)
  } catch (error) {
    return { ...base, error: `output is not a readable zip: ${describeError(error)}` }
  }

  for (const [path, content] of before) {
    if (isVolatile(path)) continue
    const other = after.get(path)
    if (!other) base.partsLost.push(path)
    else if (!bytesEqual(content, other)) base.partsChanged.push(path)
  }
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
