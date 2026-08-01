import type { RichText, RichTextRun } from './cell-input.js'
import type { XlsxErrorContext } from './errors.js'
import { type FontFormat, fontChildren, parseFont, withNamespacePrefix } from './styles-writer.js'

// A run's text lands in the sheet, so it escapes exactly as a cell value does,
// the CR fold included; styles-writer's escapeXml leaves CR alone.
const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\r/g, '&#13;')

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

/** The whole rich string flattened, as `value` reports it. */
export const flattenRuns = (runs: readonly RichTextRun[]): string =>
  runs.map((run) => run.text).join('')

/**
 * A rich string as an inline `<is>` element, each run carrying its own `rPr`.
 * Written inline rather than into the shared table, so no entry is deduped onto
 * a cell that never asked for the formatting. `prefix` matches the sheet's.
 */
export function buildRichInline(
  runs: readonly RichTextRun[],
  prefix: string,
  location: XlsxErrorContext = {},
): string {
  const body = runs
    .map((run) => {
      const rPr =
        run.font === undefined ? '' : `<rPr>${fontChildren(run.font, 'rFont', location)}</rPr>`
      const space = run.text === run.text.trim() ? '' : ' xml:space="preserve"'
      return `<r>${rPr}<t${space}>${escapeXml(run.text)}</t></r>`
    })
    .join('')
  return withNamespacePrefix(`<is>${body}</is>`, prefix)
}
