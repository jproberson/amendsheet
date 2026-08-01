import type { RichText, RichTextRun } from './cell-input.js'
import { type FontFormat, parseFont } from './styles-writer.js'

/** The font a run's `rPr` sets, or undefined when the run carries none and so
 * draws in the cell's own font. An empty `rPr` reads as no font, not `{}`. */
export function runFont(rPr: string): FontFormat | undefined {
  const font = parseFont(rPr)
  return Object.keys(font).length === 0 ? undefined : font
}

/**
 * A string built from runs, or undefined when it is not meaningfully rich: no
 * runs at all (a bare `<t>`), or a single run in the cell's own font, both of
 * which `value` already carries as plain text.
 */
export function richTextOf(runs: readonly RichTextRun[]): RichText | undefined {
  if (runs.length === 0) return undefined
  if (runs.length === 1 && runs[0]?.font === undefined) return undefined
  return { runs }
}
