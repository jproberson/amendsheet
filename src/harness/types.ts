/**
 * A normalized cell value, independent of any library's own representation.
 * Adapters translate their native types into this shape so different
 * libraries can be compared against each other.
 */
export type CellType = 'number' | 'text' | 'date' | 'boolean' | 'formula' | 'error'

export interface CellValue {
  type: CellType
  value: string | number | boolean | null
  /** Source expression, when type is 'formula'. */
  formula?: string
  /**
   * Fingerprint of the formatting this cell resolves to. Comparing this
   * per-cell measures what the user sees, unlike counting style-registry
   * entries, which a writer may legitimately dedupe or prune.
   */
  style?: string
}

export interface SheetValues {
  name: string
  /** Keyed by cell address, e.g. "B7" */
  cells: Map<string, CellValue>
}

/**
 * The contract a library must satisfy to be measured by the harness.
 * Implement this for a new library to benchmark it against the incumbents.
 */
export interface Adapter {
  name: string
  /** Read the file and write it back out, changing nothing. */
  roundTrip(bytes: Uint8Array): Promise<Uint8Array>
  /**
   * Read the file, change one cell, and write it back. Optional, because the
   * measurement only makes sense for a library that can edit in place.
   */
  edit?(bytes: Uint8Array): Promise<Uint8Array>
  /** Extract normalized values for comparison. */
  values(bytes: Uint8Array): Promise<SheetValues[]>
}

export type Parts = Map<string, Uint8Array>

export interface RoundTripResult {
  file: string
  ok: boolean
  /** Populated when the library threw instead of producing output. */
  error?: string
  partsLost: string[]
  partsAdded: string[]
  partsChanged: string[]
  featureLoss: FeatureDelta[]
  cellsLost: number
  cellsChanged: number
  cellsTotal: number
  bytesIn: number
  bytesOut: number
}

export interface FeatureDelta {
  feature: string
  before: number
  after: number
}
