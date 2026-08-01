import { XlsxError } from './errors.js'
import {
  type HeaderFooter,
  type PageBreaks,
  type PageMargins,
  type PageSetup,
  type PrintOptions,
  readHeaderFooter,
  readPageBreaks,
  readPageMargins,
  readPageSetup,
  readPrintOptions,
  withColumnBreaks,
  withHeaderFooter,
  withPageMargins,
  withPageSetup,
  withPrintOptions,
  withRowBreaks,
} from './page.js'
import type { SheetLocation } from './cell-input.js'
import { LAST_COLUMN, LAST_ROW, columnToIndex, indexToColumn } from './reference.js'

/**
 * The pending print settings for one workbook — page setup, margins, header and
 * footer, and manual breaks — per sheet path. Value checks (a scale in range, a
 * margin at least zero, a row or column that can begin a page), the merge with
 * what the file already carries, and the write into the sheet's XML all live
 * here. `document.ts` keeps only the sheet-in-package guard, which needs the
 * sheet's bytes.
 */
export interface PageStore {
  setPageSetup(path: string, setup: PageSetup, at: SheetLocation): void
  setPageMargins(path: string, margins: PageMargins, at: SheetLocation): void
  setPrintOptions(path: string, options: PrintOptions): void
  setHeaderFooter(path: string, headerFooter: HeaderFooter): void
  addRowBreak(path: string, row: number, at: SheetLocation): void
  addColumnBreak(path: string, column: string, at: SheetLocation): void
  mergedSetup(path: string, sheetBytes: Uint8Array | undefined): PageSetup
  mergedMargins(path: string, sheetBytes: Uint8Array | undefined): PageMargins
  mergedPrintOptions(
    path: string,
    sheetBytes: Uint8Array | undefined,
  ): { gridlines: boolean; headings: boolean }
  mergedHeaderFooter(path: string, sheetBytes: Uint8Array | undefined): HeaderFooter
  mergedBreaks(path: string, sheetBytes: Uint8Array | undefined): PageBreaks
  /** The sheet paths carrying any pending print edit, for the write pass. */
  paths(): Iterable<string>
  hasPending(): boolean
  /** The sheet's XML with every pending print edit for `path` applied, in the
   * schema order margins, setup, header/footer, row breaks, column breaks. */
  apply(sheetXml: string, path: string): string
}

export function createPageStore(): PageStore {
  const setup = new Map<string, PageSetup>()
  const margins = new Map<string, PageMargins>()
  const printOptions = new Map<string, PrintOptions>()
  const headerFooter = new Map<string, HeaderFooter>()
  const rowBreaks = new Map<string, Set<number>>()
  const columnBreaks = new Map<string, Set<number>>()

  return {
    setPageSetup(path, next, at) {
      if (
        next.orientation !== undefined &&
        next.orientation !== 'portrait' &&
        next.orientation !== 'landscape'
      ) {
        throw new XlsxError(
          'unwritable-value',
          `Orientation must be "portrait" or "landscape", not "${next.orientation}"`,
          { ...at },
        )
      }
      if (
        next.scale !== undefined &&
        (!Number.isInteger(next.scale) || next.scale < 10 || next.scale > 400)
      ) {
        throw new XlsxError(
          'unwritable-value',
          `Print scale ${next.scale} is not a whole percentage between 10 and 400`,
          { ...at },
        )
      }
      setup.set(path, { ...setup.get(path), ...next })
    },
    setPageMargins(path, next, at) {
      for (const side of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) {
        const value = next[side]
        if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
          throw new XlsxError('unwritable-value', `Margin ${side} ${value} is not zero or more`, {
            ...at,
          })
        }
      }
      margins.set(path, { ...margins.get(path), ...next })
    },
    setPrintOptions(path, next) {
      printOptions.set(path, { ...printOptions.get(path), ...next })
    },
    setHeaderFooter(path, next) {
      headerFooter.set(path, { ...headerFooter.get(path), ...next })
    },
    addRowBreak(path, row, at) {
      if (!Number.isInteger(row) || row < 2 || row > LAST_ROW) {
        throw new XlsxError('bad-reference', `Row ${row} cannot begin a page`, {
          ...at,
          reference: String(row),
        })
      }
      const breaks = rowBreaks.get(path) ?? new Set<number>()
      breaks.add(row - 1)
      rowBreaks.set(path, breaks)
    },
    addColumnBreak(path, column, at) {
      const index = columnToIndex(column)
      if (index < 2 || index > LAST_COLUMN) {
        throw new XlsxError('bad-reference', `Column ${column} cannot begin a page`, {
          ...at,
          reference: column,
        })
      }
      const breaks = columnBreaks.get(path) ?? new Set<number>()
      breaks.add(index - 1)
      columnBreaks.set(path, breaks)
    },
    mergedSetup(path, sheetBytes) {
      const file = sheetBytes === undefined ? {} : readPageSetup(sheetBytes)
      return { ...file, ...setup.get(path) }
    },
    mergedMargins(path, sheetBytes) {
      const file = sheetBytes === undefined ? {} : readPageMargins(sheetBytes)
      return { ...file, ...margins.get(path) }
    },
    mergedPrintOptions(path, sheetBytes) {
      const file =
        sheetBytes === undefined
          ? { gridlines: false, headings: false }
          : readPrintOptions(sheetBytes)
      const pending = printOptions.get(path)
      return {
        gridlines: pending?.gridlines ?? file.gridlines,
        headings: pending?.headings ?? file.headings,
      }
    },
    mergedHeaderFooter(path, sheetBytes) {
      const file = sheetBytes === undefined ? {} : readHeaderFooter(sheetBytes)
      return { ...file, ...headerFooter.get(path) }
    },
    mergedBreaks(path, sheetBytes) {
      const file = sheetBytes === undefined ? { rows: [], columns: [] } : readPageBreaks(sheetBytes)
      const rows = new Set<number>(file.rows)
      for (const id of rowBreaks.get(path) ?? []) rows.add(id + 1)
      const columns = new Set<string>(file.columns)
      for (const id of columnBreaks.get(path) ?? []) columns.add(indexToColumn(id + 1))
      return {
        rows: [...rows].sort((a, b) => a - b),
        columns: [...columns].sort((a, b) => columnToIndex(a) - columnToIndex(b)),
      }
    },
    paths() {
      return new Set([
        ...margins.keys(),
        ...setup.keys(),
        ...printOptions.keys(),
        ...headerFooter.keys(),
        ...rowBreaks.keys(),
        ...columnBreaks.keys(),
      ])
    },
    hasPending() {
      return (
        setup.size > 0 ||
        margins.size > 0 ||
        printOptions.size > 0 ||
        headerFooter.size > 0 ||
        rowBreaks.size > 0 ||
        columnBreaks.size > 0
      )
    },
    apply(sheetXml, path) {
      let xml = sheetXml
      const printOptionsFor = printOptions.get(path)
      if (printOptionsFor !== undefined) xml = withPrintOptions(xml, printOptionsFor)
      const marginsFor = margins.get(path)
      if (marginsFor !== undefined) xml = withPageMargins(xml, marginsFor)
      const setupFor = setup.get(path)
      if (setupFor !== undefined) xml = withPageSetup(xml, setupFor)
      const headerFooterFor = headerFooter.get(path)
      if (headerFooterFor !== undefined) xml = withHeaderFooter(xml, headerFooterFor)
      const rowBreaksFor = rowBreaks.get(path)
      if (rowBreaksFor !== undefined) xml = withRowBreaks(xml, [...rowBreaksFor])
      const columnBreaksFor = columnBreaks.get(path)
      if (columnBreaksFor !== undefined) xml = withColumnBreaks(xml, [...columnBreaksFor])
      return xml
    },
  }
}
