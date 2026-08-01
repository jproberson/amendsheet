import type { FontFormat } from './styles-writer.js'

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

/** One stretch of a rich string: its text and, when it carries formatting of its
 * own, the font that stretch is drawn in. A run with no `font` inherits the cell's. */
export interface RichTextRun {
  readonly text: string
  readonly font?: FontFormat
}

/** A string whose parts carry their own formatting, so one cell can mix fonts,
 * weights and colours. Written by `set`, read back on `cell.richText`. */
export interface RichText {
  readonly runs: readonly RichTextRun[]
}

export type CellInput = number | string | boolean | Date | null | FormulaInput
