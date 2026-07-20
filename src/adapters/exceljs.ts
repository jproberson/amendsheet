import ExcelJS from 'exceljs'
import { toArrayBuffer } from '../harness/zip.js'
import type { Adapter, CellValue, SheetValues } from '../harness/types.js'

/**
 * ExcelJS declares `interface Buffer extends ArrayBuffer`, which is untrue of
 * Node's Buffer. Its loader passes the argument straight to JSZip, which
 * accepts an ArrayBuffer, so handing it one satisfies both the declaration and
 * the runtime without an assertion.
 */
async function loadWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(toArrayBuffer(bytes))
  return workbook
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function hasFormula(
  value: Record<string, unknown>,
): value is { formula?: string; sharedFormula?: string; result?: unknown } {
  return typeof value.formula === 'string' || typeof value.sharedFormula === 'string'
}

function hasRichText(
  value: Record<string, unknown>,
): value is { richText: Array<{ text: string }> } {
  return (
    Array.isArray(value.richText) &&
    value.richText.every((run) => isObject(run) && typeof run.text === 'string')
  )
}

function hasHyperlink(
  value: Record<string, unknown>,
): value is { hyperlink: string; text?: string } {
  return typeof value.hyperlink === 'string'
}

function hasError(value: Record<string, unknown>): value is { error: string } {
  return typeof value.error === 'string'
}

/**
 * Translate ExcelJS's union-typed cell values into the harness's normal form.
 * ExcelJS represents formulas, hyperlinks, rich text and errors as objects,
 * so a naive comparison would report spurious differences.
 */
function normalizeCellValue(raw: unknown): CellValue | null {
  if (raw === null || raw === undefined || raw === '') return null

  if (raw instanceof Date) return { type: 'date', value: raw.toISOString() }
  if (typeof raw === 'number') return { type: 'number', value: raw }
  if (typeof raw === 'boolean') return { type: 'boolean', value: raw }
  if (typeof raw === 'string') return { type: 'text', value: raw }
  if (!isObject(raw)) return null

  if (hasFormula(raw)) {
    const source = raw.formula ?? raw.sharedFormula ?? ''
    const cached = normalizeCellValue(raw.result)
    return { type: 'formula', value: cached?.value ?? null, formula: source }
  }
  if (hasRichText(raw)) {
    return { type: 'text', value: raw.richText.map((run) => run.text).join('') }
  }
  if (hasHyperlink(raw)) {
    return { type: 'text', value: raw.text ?? raw.hyperlink }
  }
  if (hasError(raw)) {
    return { type: 'error', value: raw.error }
  }
  return { type: 'text', value: JSON.stringify(raw) }
}

/**
 * Reduce a cell's resolved formatting to a stable string. Only properties a
 * user would notice are included, so cosmetic differences in how a writer
 * organizes its style tables do not count as loss.
 */
function styleFingerprint(cell: ExcelJS.Cell): string | undefined {
  const parts: string[] = []
  const { font, fill, border, alignment, numFmt } = cell

  if (numFmt) parts.push(`format=${numFmt}`)

  if (font) {
    const traits = [
      font.bold ? 'bold' : '',
      font.italic ? 'italic' : '',
      font.underline ? 'underline' : '',
      font.size ? `size${font.size}` : '',
      font.name ?? '',
      typeof font.color?.argb === 'string' ? font.color.argb : '',
    ].filter(Boolean)
    if (traits.length) parts.push(`font=${traits.join(',')}`)
  }

  if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
    const foreground = typeof fill.fgColor?.argb === 'string' ? fill.fgColor.argb : ''
    parts.push(`fill=${fill.pattern}:${foreground}`)
  }

  if (border) {
    const edges: string[] = []
    for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
      const style = border[edge]?.style
      if (style) edges.push(`${edge}:${style}`)
    }
    if (edges.length) parts.push(`border=${edges.join(',')}`)
  }

  if (alignment?.horizontal || alignment?.vertical) {
    parts.push(`align=${alignment.horizontal ?? ''}:${alignment.vertical ?? ''}`)
  }

  return parts.length ? parts.join('|') : undefined
}

export const exceljsAdapter: Adapter = {
  name: 'exceljs@4.4.0',

  async roundTrip(bytes) {
    const workbook = await loadWorkbook(bytes)
    const written = await workbook.xlsx.writeBuffer()
    return new Uint8Array(written)
  },

  async values(bytes) {
    const workbook = await loadWorkbook(bytes)
    const sheets: SheetValues[] = []

    workbook.eachSheet((worksheet) => {
      const cells = new Map<string, CellValue>()
      worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const value = normalizeCellValue(cell.value)
          if (value === null) return
          const style = styleFingerprint(cell)
          cells.set(cell.address, style ? { ...value, style } : value)
        })
      })
      sheets.push({ name: worksheet.name, cells })
    })

    return sheets
  },
}
