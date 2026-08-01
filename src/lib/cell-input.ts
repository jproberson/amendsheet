/**
 * Where a write-path error points, past the cell reference the throw site
 * already knows. `set()` fills it in from the sheet the edit is on, so a
 * refusal names the sheet and part and not just the cell.
 */
export interface SheetLocation {
  readonly sheet?: string
  readonly part?: string
}

/** An expression without the leading `=`, so text starting with `=` stays text. */
export interface FormulaInput {
  readonly formula: string
}

export type CellInput = number | string | boolean | Date | null | FormulaInput
