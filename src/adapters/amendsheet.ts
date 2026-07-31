import { readWorkbook } from '../lib/document.js'
import type { CellFormula, CellValue as DocumentValue } from '../lib/public-types.js'
import type { Adapter, CellValue, SheetValues } from '../harness/types.js'

/** A dependent is identified by its master, which is what exceljs reports too. */
function toHarnessFormula(formula: CellFormula): string {
  return formula.kind === 'expression' ? formula.expression : (formula.master ?? '')
}

function toHarnessValue(value: DocumentValue): CellValue | null {
  switch (value.kind) {
    case 'empty':
      return null
    case 'date':
      return { type: 'date', value: value.value.toISOString() }
    case 'number':
      return { type: 'number', value: value.value }
    case 'text':
      return value.value === '' ? null : { type: 'text', value: value.value }
    case 'boolean':
      return { type: 'boolean', value: value.value }
    case 'error':
      return { type: 'error', value: value.value }
  }
}

/**
 * Measures this library with the same rig used on the incumbents, so the
 * comparison runs over identical files and identical assertions.
 */
export const amendsheetAdapter: Adapter = {
  name: 'amendsheet (this library)',

  async roundTrip(bytes) {
    return readWorkbook(bytes).toBytes()
  },

  async edit(bytes) {
    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    if (sheet === undefined) return workbook.toBytes()

    // Past the last row in use, so nothing already in the file is overwritten
    // and every existing cell must still come back unchanged.
    let lastRow = 0
    for (const cell of sheet.cells()) lastRow = Math.max(lastRow, cell.address.row)
    sheet.set(`A${lastRow + 1}`, 'amendsheet harness')

    return workbook.toBytes()
  },

  async values(bytes) {
    const workbook = readWorkbook(bytes)
    const sheets: SheetValues[] = []

    for (const worksheet of workbook.sheets) {
      const cells = new Map<string, CellValue>()

      for (const cell of worksheet.cells()) {
        const value = toHarnessValue(cell.value)
        if (value === null) continue

        cells.set(cell.reference, {
          ...value,
          ...(cell.formula === undefined ? {} : { formula: toHarnessFormula(cell.formula) }),
          ...(cell.numberFormat === undefined ? {} : { style: `format=${cell.numberFormat}` }),
        })
      }

      sheets.push({ name: worksheet.name, cells })
    }

    return sheets
  },
}
