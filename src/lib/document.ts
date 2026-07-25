import {
  type AddedSheet,
  EMPTY_SHEET_XML,
  checkSheetName,
  withSheetContentTypes,
  withSheetRelationships,
  withSheetRemoved,
  withSheetRenamed,
  withSheetsAdded,
} from './add-sheet.js'
import { blankWorkbookBytes } from './blank.js'
import { checkDefinedName, readDefinedNames, withDefinedNames } from './defined-names.js'
import { type HyperlinkEntry, withHyperlinkRelationships, withHyperlinks } from './hyperlinks.js'
import { type Container, decodeXmlPart } from './container.js'
import { XlsxError } from './errors.js'
import { LAST_SERIAL, dateToSerial, parseIsoDate, serialToDate } from './date.js'
import {
  type CellInput,
  checkProtection,
  checkWritable,
  mergeAnchorFor,
  mergeRangeReference,
  mergeRefusal,
  patchSheet,
  indexSheet,
  readSheetProtection,
  sharedFormulaRefusal,
  type SheetIndex,
  type SheetLocation,
  type SheetProtection,
} from './patch.js'
import {
  type CellAddress,
  LAST_COLUMN,
  LAST_ROW,
  canonicalReference,
  columnToIndex,
  formatReference,
  parseReference,
  parseWritableReference,
} from './reference.js'
import { type RawCell, readSheet } from './sheet.js'
import { appendSharedStrings, readSharedStrings } from './shared-strings.js'
import { extendTables } from './tables.js'
import {
  type Alignment,
  type BorderFormat,
  type CellFormatting,
  type CellProtection,
  type DateStyle,
  type FillFormat,
  type FontFormat,
  checkStyleOptions,
  ensureAlignmentStyle,
  ensureBorderStyle,
  ensureDateStyle,
  ensureFillStyle,
  ensureFontStyle,
  ensureNumberFormat,
  ensureProtectionStyle,
  readFormatting,
} from './styles-writer.js'
import { type ShiftSpec, shiftFormula } from './shift.js'
import { shiftForeignFormulas, shiftSheet } from './shift-sheet.js'
import { type Styles, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import {
  deletionDamage,
  highestColumn,
  highestRow,
  relationshipsPathFor,
  unshiftablePart,
  withRecalculation,
  withoutOverride,
  withoutRelationship,
} from './workbook-parts.js'
import { type SheetRef, type SheetState, readWorkbookPart } from './workbook.js'

export type CellValue =
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly value: string }
  | { readonly kind: 'empty' }
  /**
   * The stored number is kept so a date can be written back unchanged. Which
   * formats count as dates is a heuristic and may change in any release, so a
   * cell can move between `number` and `date`; `serial` and `value` are the
   * same double, and reading whichever is there survives that.
   */
  | { readonly kind: 'date'; readonly value: Date; readonly serial: number }

/**
 * A shared formula stores its source once, on the cell that owns the range.
 * The dependents carry a cached value and no source of their own, which is why
 * this is a union rather than a string that is sometimes empty.
 */
export type CellFormula =
  /** Without the leading `=`. */
  | { readonly kind: 'expression'; readonly expression: string }
  /** `master` is absent when the sheet holds no cell owning that group. */
  | { readonly kind: 'shared'; readonly master?: string }

export interface Cell {
  readonly address: CellAddress
  /**
   * Canonical, so it always equals `formatReference(cell.address)`. The one
   * exception is an address no column letter can name, which a lenient read
   * accepts and which keeps the spelling the file gave it.
   */
  readonly reference: string
  readonly value: CellValue
  readonly formula?: CellFormula
  readonly numberFormat?: string
  /** Absent when the cell uses the default font, so carries none of its own. */
  readonly font?: FontFormat
  /** Absent unless the cell has a solid fill. */
  readonly fill?: FillFormat
  /** Absent unless at least one side has a border. */
  readonly border?: BorderFormat
  /** Absent unless the cell sets an alignment of its own. */
  readonly alignment?: Alignment
  /** Absent unless the cell sets a `locked` or `hidden` protection of its own. */
  readonly protection?: CellProtection
}

/**
 * Where a cell links to: a URL out of the package, or a `location` within the
 * workbook — a cell reference like `Sheet2!A1` or a defined name. `tooltip` is the
 * hover text.
 */
export type Hyperlink =
  | { readonly url: string; readonly tooltip?: string }
  | { readonly location: string; readonly tooltip?: string }

export interface Worksheet {
  readonly name: string
  /** Renames the sheet. The name follows the same rules as `addSheet`. */
  rename(name: string): void
  /** Removes the sheet from the workbook. A workbook must keep at least one. */
  remove(): void
  readonly state: SheetState
  /**
   * As the workbook part spells it, so a defined name or a part this library
   * does not interpret can be matched against the sheet it refers to.
   */
  readonly sheetId: string
  /**
   * The worksheet protection in force, in the shape `protect()` takes, or
   * undefined when the sheet is not protected. Reflects a pending `protect()` or
   * `unprotect()` as well as what the file was read with.
   */
  readonly protection?: SheetProtection
  /**
   * Every cell the sheet stores. A cell that was cleared, or that carries only
   * formatting, is still stored, and arrives with a value of `kind: 'empty'`.
   *
   * Each call re-reads the sheet, so a call per cell is quadratic.
   */
  cells(): Iterable<Cell>
  /** Undefined when the sheet stores nothing at that reference. */
  cell(reference: string): Cell | undefined
  /**
   * Visible to `cells()` and `cell()` immediately, written by `toBytes()`.
   * A `numberFormat` is a format code such as `"$"#,##0.00`; without one the
   * cell keeps the formatting it already had.
   *
   * Throws `XlsxError` with code `unwritable-value` for a value the format
   * cannot hold, and records nothing when it does, so the rest of a batch of
   * edits still writes.
   */
  set(reference: string, value: CellInput, options?: SetOptions): void
  /**
   * Applies formatting to a cell without changing its value or formula, so a
   * formula cell can be restyled without losing its expression. A cell that is
   * not there yet is created empty with the formatting. Unlike `set`, this is
   * refused by nothing: restyling a shared-formula master or a merged cell is
   * safe, since only the cell format changes.
   */
  format(reference: string, options: SetOptions): void
  /**
   * Turns on worksheet protection, which is what makes a cell's `locked` and
   * `hidden` flags bite. Without `options` it matches Excel's Protect Sheet
   * default; `options` names the actions that stay permitted. Replaces any
   * protection the sheet already declared. Passwords are not written.
   */
  protect(options?: SheetProtection): void
  /** Removes worksheet protection, if the sheet had any. */
  unprotect(): void
  /**
   * Merges a rectangular range like `A1:B2`, joining any merges the sheet
   * already has. Excel shows only the top-left cell's value; the others keep
   * whatever they hold, since a merge does not clear them. Refuses a range that
   * is not two references either side of a colon.
   */
  merge(range: string): void
  /** Sets the sheet's auto-filter over a range, replacing any it already has. */
  autoFilter(range: string): void
  /**
   * Freezes the rows above and the columns left of `cell`, so they stay in view
   * when the sheet is scrolled. `freeze('B2')` freezes row 1 and column A.
   */
  freeze(cell: string): void
  /**
   * Sets a row's height in points, marking it a custom height so a reader keeps
   * it. Refuses a row number below 1 or a height that is not a finite number at
   * least zero.
   */
  setRowHeight(row: number, height: number): void
  /**
   * Sets a column's width, in the units Excel shows, splitting a `cols` range
   * that spans more than the column so the rest keeps its own width. The column
   * is a letter like `A`. Refuses a width that is not a finite number at least
   * zero.
   */
  setColumnWidth(column: string, width: number): void
  /** Hides a row, keeping any height it also has. The row is one-based. */
  hideRow(row: number): void
  /** Hides a column, keeping any width it also has. The column is a letter. */
  hideColumn(column: string): void
  /**
   * Inserts `count` blank rows before row `before`, pushing the rows at and below
   * it down. References that point into the moved rows — formulas anywhere in the
   * workbook, merges, the dimension, filters, conditional formats and defined
   * names — move with them. Cell edits made in this session land first, so they
   * ride along too. Refused when the sheet carries a table, drawing, pivot table
   * or comment, whose pinned positions this does not yet adjust, or when the
   * shift would push a row off the sheet.
   */
  insertRows(before: number, count?: number): void
  /**
   * Inserts `count` blank columns before column `before` — a letter like `C` —
   * pushing the columns at and to the right of it over. References that point into
   * the moved columns move with them, on the same terms as `insertRows`, a pinned
   * part and an off-the-sheet shift included.
   */
  insertColumns(before: string, count?: number): void
  /**
   * Deletes `count` rows from row `from`, pulling the rows below up over them. A
   * reference into a deleted row becomes #REF! where a formula named the cell, and
   * shrinks where a range only overlapped. Refused when a whole merged range,
   * filter, format or shared formula sat inside the deletion, none of which can
   * survive as #REF!, or when the sheet carries a table, drawing, pivot table or
   * comment.
   */
  deleteRows(from: number, count?: number): void
  /**
   * Deletes `count` columns from column `from` — a letter like `C` — pulling the
   * columns to their right in over them, on the same terms as `deleteRows`.
   */
  deleteColumns(from: string, count?: number): void
  /**
   * Links a cell to a URL or a place in the workbook, replacing any link the cell
   * already has. An external URL is written through a worksheet relationship; a
   * `location` is written inline. Visible in the file after `toBytes()`.
   */
  link(reference: string, target: Hyperlink): void
}

export interface SetOptions {
  /** A number format code, applied to the cell being written. */
  readonly numberFormat?: string
  /** Font to apply, merged onto the font the cell already carries. */
  readonly font?: FontFormat
  /** Solid fill colour for the cell's background. */
  readonly fill?: FillFormat
  /** Borders to apply, merged onto the sides the cell already has. */
  readonly border?: BorderFormat
  /** Alignment to apply, merged onto the alignment the cell already has. */
  readonly alignment?: Alignment
  /** Protection to apply, merged onto the protection the cell already has. */
  readonly protection?: CellProtection
}

export interface Workbook {
  readonly sheets: readonly Worksheet[]
  /** Undefined when no sheet has that name. Names are compared exactly. */
  sheet(name: string): Worksheet | undefined
  /**
   * Adds an empty worksheet and returns it, ready to fill. The name must be one
   * Excel accepts — up to 31 characters, none of `: \ / ? * [ ]`, and not one a
   * sheet already uses. Written into the workbook by `toBytes()`.
   */
  addSheet(name: string): Worksheet
  /**
   * The workbook's global named ranges, each mapped to what it refers to (a
   * formula like `Sheet1!$A$1:$B$2`). Reflects a pending `defineName`.
   */
  readonly definedNames: ReadonlyMap<string, string>
  /**
   * Defines a global named range, replacing one of the same name. The name
   * follows Excel's rules — a letter, underscore or backslash then letters,
   * digits, periods or underscores, no spaces. `refersTo` is a formula.
   */
  defineName(name: string, refersTo: string): void
  /** Which year serials count from. A 1904 workbook is 1462 days behind. */
  readonly epoch: 1900 | 1904
  /** Parts that were never interpreted are written exactly as they were read. */
  toBytes(): Uint8Array
}

const EMPTY_STYLES: Styles = { numberFormats: new Map(), cellFormats: [] }
const EMPTY_EDITS: ReadonlyMap<string, CellInput> = new Map()

const CALCULATION_CHAIN = 'xl/calcChain.xml'
const CONTENT_TYPES = '[Content_Types].xml'

function partText(container: Container, path: string): string | undefined {
  const bytes = container.parts.get(path)
  if (bytes === undefined) return undefined
  return decodeXmlPart(bytes, path)
}

function toCellValue(raw: RawCell, styles: Styles, date1904: boolean): CellValue {
  const value = raw.value

  if (value.kind === 'date') {
    const parsed = parseIsoDate(value.value)
    if (parsed === undefined) return { kind: 'text', value: value.value }
    return { kind: 'date', value: parsed, serial: dateToSerial(parsed, date1904) }
  }

  if (value.kind === 'number' && isDateFormat(styles, raw.styleIndex)) {
    const serial = value.value
    // A serial outside the range dates cover stays the number it is. Excel
    // shows such a cell as ###, so throwing here would make a legal file
    // unreadable.
    if (serial >= 0 && serial <= LAST_SERIAL) {
      return { kind: 'date', value: serialToDate(serial, date1904), serial }
    }
  }

  return value
}

/** Where each shared group's source lives, filled in as the sheet is read. */
type SharedMasters = Map<string, string>

function toFormula(raw: RawCell, masters: SharedMasters | undefined): CellFormula | undefined {
  if (raw.formula === undefined) return undefined
  if (raw.sharedIndex === undefined || raw.ownsSharedRange === true) {
    return { kind: 'expression', expression: raw.formula }
  }

  const master = masters?.get(raw.sharedIndex)
  return master === undefined ? { kind: 'shared' } : { kind: 'shared', master }
}

function toCell(
  raw: RawCell,
  styles: Styles,
  formatting: CellFormatting,
  date1904: boolean,
  masters?: SharedMasters,
): Cell {
  const numberFormat = numberFormatOf(styles, raw.styleIndex)
  const value = toCellValue(raw, styles, date1904)
  const formula = toFormula(raw, masters)

  return {
    address: raw.address,
    // Not the file's spelling: $A$1 and a1 are the same cell, and a caller that
    // cross-references this against set() or formatReference needs one answer.
    reference: canonicalReference(raw.address) ?? raw.reference,
    value,
    ...(formula === undefined ? {} : { formula }),
    ...(numberFormat === undefined ? {} : { numberFormat }),
    ...formatting,
  }
}

/**
 * A blank workbook with one empty sheet — named `Sheet1` unless a name is given —
 * to fill with `set()` and the rest of the edit API, then write with `toBytes()`.
 * Creating is amending an empty file, so it comes back through the same path a
 * read one does. The name must be one Excel accepts, the way `addSheet()` asks.
 */
export function createWorkbook(sheetName = 'Sheet1'): Workbook {
  checkSheetName(sheetName, [])
  return readWorkbook(blankWorkbookBytes(sheetName))
}

export function readWorkbook(bytes: Uint8Array): Workbook {
  const part = readWorkbookPart(bytes)
  const { container, date1904 } = part

  const stylesXml = partText(container, 'xl/styles.xml')
  const styles = stylesXml === undefined ? EMPTY_STYLES : readStyles(stylesXml)

  const stringsXml = partText(container, 'xl/sharedStrings.xml')
  const sharedStrings = stringsXml === undefined ? [] : readSharedStrings(stringsXml)

  const edits = new Map<string, Map<string, CellInput>>()

  // Which cell format each edited cell lands on is decided by set(), not by
  // toBytes(). Choosing a style index is what makes a number a date, so it is a
  // decision about what a cell MEANS; leaving it to write time gave the read
  // path its own copy of the decision, and the two drifted.
  const styleOverrides = new Map<string, Map<string, number>>()
  const sheetProtections = new Map<string, SheetProtection | 'remove'>()
  const sheetMerges = new Map<string, string[]>()
  const sheetRowHeights = new Map<string, Map<number, number>>()
  const sheetColumnWidths = new Map<string, Map<number, number>>()
  const sheetAutoFilters = new Map<string, string>()
  const sheetFreezes = new Map<string, string>()
  const sheetHiddenRows = new Map<string, Set<number>>()
  const sheetHiddenColumns = new Map<string, Set<number>>()
  // The per-sheet maps patchSheet applies in one rewrite. Both the "anything
  // pending?" check and the set of sheets to rewrite read this list, so a new
  // kind of sheet edit is registered in one place rather than two enumerations
  // that have to be kept in step.
  const patchInputs: ReadonlyArray<ReadonlyMap<string, unknown>> = [
    edits,
    styleOverrides,
    sheetProtections,
    sheetMerges,
    sheetRowHeights,
    sheetColumnWidths,
    sheetAutoFilters,
    sheetFreezes,
    sheetHiddenRows,
    sheetHiddenColumns,
  ]
  const fileNames = readDefinedNames(partText(container, part.path) ?? '')
  const pendingNames = new Map<string, string>()
  const sheetHyperlinks = new Map<string, Map<string, Hyperlink>>()
  // Row and column inserts and deletes, in call order. Each names the sheet it
  // was called on; toBytes applies them after the per-sheet patch so an edit made
  // this session lands in the old grid and then moves with the shift.
  const lineOps: { readonly path: string; readonly spec: ShiftSpec }[] = []
  let workingStyles = stylesXml
  let parsedStyles = styles
  let parsedFrom = stylesXml

  const stylesNow = (): Styles => {
    if (workingStyles !== undefined && workingStyles !== parsedFrom) {
      parsedStyles = readStyles(workingStyles)
      parsedFrom = workingStyles
    }
    return parsedStyles
  }

  // Parsed lazily, and only for the font/fill/border a read exposes, so a
  // workbook that is never read for formatting never pays to resolve it.
  let parsedFormatting: readonly CellFormatting[] | undefined
  let formattingFrom: string | undefined
  const formattingFor = (styleIndex: number | undefined): CellFormatting => {
    if (styleIndex === undefined || workingStyles === undefined) return {}
    if (parsedFormatting === undefined || workingStyles !== formattingFrom) {
      parsedFormatting = readFormatting(workingStyles)
      formattingFrom = workingStyles
    }
    return parsedFormatting[styleIndex] ?? {}
  }

  // Sheets added since the read, path -> its empty worksheet bytes. They flow
  // through the same edit machinery as a sheet the file already had; toBytes()
  // writes the new part and wires it into the workbook, its rels and the types.
  const addedSheets = new Map<string, Uint8Array>()
  const addedRefs: AddedSheet[] = []
  // A sheet renamed since the read, path -> its new name. toBytes() rewrites the
  // workbook's <sheet> for a sheet the file had, and the wiring for an added one.
  const renames = new Map<string, string>()
  // Sheets removed since the read, by path. A sheet the file had is unwired and
  // its part dropped; an added one is simply never written.
  const removed = new Set<string>()

  const makeWorksheet = (reference: SheetRef): Worksheet => {
    const sheetBytes = container.parts.get(reference.path) ?? addedSheets.get(reference.path)
    const at: SheetLocation = { sheet: reference.name, part: reference.path }

    const patched = (): Uint8Array | undefined => {
      if (sheetBytes === undefined) return undefined
      const pending = edits.get(reference.path)
      const overrides = styleOverrides.get(reference.path)
      const protection = sheetProtections.get(reference.path)
      const merges = sheetMerges.get(reference.path)
      const rowHeights = sheetRowHeights.get(reference.path)
      const columnWidths = sheetColumnWidths.get(reference.path)
      // A format() with no set() leaves overrides but no pending edit.
      if (
        pending === undefined &&
        overrides === undefined &&
        protection === undefined &&
        merges === undefined &&
        rowHeights === undefined &&
        columnWidths === undefined
      ) {
        return sheetBytes
      }
      return patchSheet(sheetBytes, pending ?? EMPTY_EDITS, date1904, undefined, overrides, at, {
        protection,
        merges,
        rowHeights,
        columnWidths,
      })
    }

    function* readCells(source?: Uint8Array): Generator<Cell> {
      const bytes = source ?? patched()
      if (bytes === undefined) return

      // The cell owning a shared range comes before its dependents, so filling
      // this in as the sheet streams resolves every one of them.
      const masters: SharedMasters = new Map()

      for (const raw of readSheet(bytes, sharedStrings, at)) {
        if (raw.ownsSharedRange === true && raw.sharedIndex !== undefined) {
          masters.set(raw.sharedIndex, canonicalReference(raw.address) ?? raw.reference)
        }
        yield toCell(raw, stylesNow(), formattingFor(raw.styleIndex), date1904, masters)
      }
    }

    // The sheet as it was read: which style each cell carried, and which cells
    // define a shared formula. Built once, on the first write, because every
    // set() needs both and parsing the sheet twice per sheet showed up in the
    // bench.
    let index: SheetIndex | undefined
    const indexed = (): SheetIndex | undefined => {
      if (sheetBytes === undefined) return undefined
      index ??= indexSheet(sheetBytes)
      return index
    }

    const styleAt = (canonical: string): number | undefined =>
      styleOverrides.get(reference.path)?.get(canonical) ?? indexed()?.styles.get(canonical)

    /**
     * What the cell becomes, resolved through the same function a read uses.
     * Predicting the value independently is how the overlay came to disagree
     * with the file it was about to write.
     */
    const predict = (canonical: string, value: CellInput, styleIndex: number | undefined): Cell => {
      const address = parseReference(canonical)
      const style = styleIndex === undefined ? {} : { styleIndex }
      const raw = { address, reference: canonical, ...style }
      const formatting = formattingFor(styleIndex)

      if (value === null) {
        return toCell({ ...raw, value: { kind: 'empty' } }, stylesNow(), formatting, date1904)
      }
      if (typeof value === 'number') {
        return toCell(
          { ...raw, value: { kind: 'number', value } },
          stylesNow(),
          formatting,
          date1904,
        )
      }
      if (typeof value === 'boolean') {
        return toCell(
          { ...raw, value: { kind: 'boolean', value } },
          stylesNow(),
          formatting,
          date1904,
        )
      }
      if (typeof value === 'string') {
        return toCell({ ...raw, value: { kind: 'text', value } }, stylesNow(), formatting, date1904)
      }
      if (value instanceof Date) {
        // Written as the serial it becomes, so the style decides how it reads.
        const serial = dateToSerial(value, date1904)
        return toCell(
          { ...raw, value: { kind: 'number', value: serial } },
          stylesNow(),
          formatting,
          date1904,
        )
      }
      // A formula is written without a cached value, so it reads back empty.
      const formula = value.formula
      return toCell(
        { ...raw, value: { kind: 'empty' }, formula },
        stylesNow(),
        formatting,
        date1904,
      )
    }

    // Built from the sheet as it was read. Every edit is in the overlay, so the
    // index never needs rebuilding.
    let byReference: Map<string, Cell> | undefined
    const overlay = new Map<string, Cell>()

    const findCell = (canonical: string): Cell | undefined => {
      const edited = overlay.get(canonical)
      if (edited !== undefined) return edited
      if (byReference === undefined) {
        byReference = new Map()
        for (const found of readCells(sheetBytes)) {
          const where = canonicalReference(found.address)
          if (where !== undefined) byReference.set(where, found)
        }
      }
      return byReference.get(canonical)
    }

    // A restyled cell keeps its value and formula; only its number format is
    // re-resolved, which can move a number to a date and back.
    const restyled = (base: Cell, styleIndex: number | undefined): Cell => {
      const styles = stylesNow()
      const numberFormat = numberFormatOf(styles, styleIndex)
      const dateFormat = isDateFormat(styles, styleIndex)
      let value = base.value
      if (value.kind === 'number' && dateFormat && value.value >= 0 && value.value <= LAST_SERIAL) {
        value = { kind: 'date', value: serialToDate(value.value, date1904), serial: value.value }
      } else if (value.kind === 'date' && !dateFormat) {
        value = { kind: 'number', value: value.serial }
      }
      return {
        address: base.address,
        reference: base.reference,
        value,
        ...(base.formula === undefined ? {} : { formula: base.formula }),
        ...(numberFormat === undefined ? {} : { numberFormat }),
        ...formattingFor(styleIndex),
      }
    }

    /**
     * A number format and a font composed into one cell format, each step landing
     * on the one before. Undefined when there is nothing to change. Refuses at the
     * call, before anything is recorded, so a rejected format queues nothing.
     */
    const resolveStyle = (
      current: number | undefined,
      value: CellInput | undefined,
      options: SetOptions | undefined,
      canonical: string,
    ): DateStyle | undefined => {
      if (workingStyles === undefined) {
        if (Object.values(options ?? {}).some((asked) => asked !== undefined)) {
          throw new XlsxError(
            'missing-part',
            `Cannot format ${canonical}: the package has no style table`,
            { part: 'xl/styles.xml', reference: canonical },
          )
        }
        return undefined
      }
      let xml = workingStyles
      let base = current
      let applied: DateStyle | undefined
      const step = (next: DateStyle) => {
        xml = next.xml
        base = next.index
        applied = next
      }
      // An asked-for format wins; a Date only gets one because without one it
      // displays as the serial number it is stored as.
      if (options?.numberFormat !== undefined)
        step(ensureNumberFormat(xml, base, options.numberFormat))
      else if (value instanceof Date) step(ensureDateStyle(xml, base))
      if (options?.font !== undefined) step(ensureFontStyle(xml, base, options.font))
      if (options?.fill !== undefined) step(ensureFillStyle(xml, base, options.fill))
      if (options?.border !== undefined) step(ensureBorderStyle(xml, base, options.border))
      if (options?.alignment !== undefined) step(ensureAlignmentStyle(xml, base, options.alignment))
      if (options?.protection !== undefined)
        step(ensureProtectionStyle(xml, base, options.protection))
      return applied
    }

    const commitStyle = (
      canonical: string,
      current: number | undefined,
      applied: DateStyle | undefined,
    ): number | undefined => {
      if (applied === undefined) return current
      workingStyles = applied.xml
      if (applied.index !== current) {
        const overrides = styleOverrides.get(reference.path) ?? new Map<string, number>()
        overrides.set(canonical, applied.index)
        styleOverrides.set(reference.path, overrides)
      }
      return applied.index
    }

    const absent = (canonical: string, verb: string): XlsxError =>
      new XlsxError(
        'missing-part',
        `Sheet ${reference.name} is not in the package, so ${canonical} cannot be ${verb}`,
        { ...at, reference: canonical },
      )

    // A part with pinned positions is preserved untouched, so an insert or delete
    // that would move the cells under it is refused until it too can be adjusted.
    const refuseUnshiftable = (action: string, where?: string): void => {
      const relationships = partText(container, relationshipsPathFor(reference.path))
      if (relationships === undefined) return
      const owns = unshiftablePart(relationships)
      if (owns === undefined) return
      throw new XlsxError(
        'unsupported-edit',
        `Sheet ${reference.name} carries ${owns}, so ${action}`,
        { ...at, ...(where === undefined ? {} : { reference: where }) },
      )
    }

    const lineSpec = (axis: 'row' | 'column', line: number, delta: number): ShiftSpec => ({
      axis,
      at: line,
      delta,
      editedSheet: reference.name,
      onCurrentSheet: true,
    })

    const checkCount = (count: number, noun: string, where?: string): void => {
      if (!Number.isInteger(count) || count < 1) {
        throw new XlsxError('unwritable-value', `${count} is not a number of ${noun}`, {
          ...at,
          ...(where === undefined ? {} : { reference: where }),
        })
      }
    }

    return {
      get name(): string {
        return renames.get(reference.path) ?? reference.name
      },
      rename(name: string): void {
        const current = renames.get(reference.path) ?? reference.name
        checkSheetName(
          name,
          sheets.map((sheet) => sheet.name).filter((other) => other !== current),
        )
        renames.set(reference.path, name)
      },
      remove(): void {
        if (sheets.length <= 1) {
          throw new XlsxError('unwritable-value', 'A workbook must keep at least one sheet', {
            ...at,
          })
        }
        const current = renames.get(reference.path) ?? reference.name
        const index = sheets.findIndex((sheet) => sheet.name === current)
        if (index !== -1) sheets.splice(index, 1)
        removed.add(reference.path)
        // An added sheet has no part in the package, so undo its registration and
        // it is simply never written; a read sheet is unwired at toBytes().
        if (addedSheets.has(reference.path)) {
          addedSheets.delete(reference.path)
          const addedIndex = addedRefs.findIndex((added) => added.reference.path === reference.path)
          if (addedIndex !== -1) addedRefs.splice(addedIndex, 1)
        }
      },
      state: reference.state,
      sheetId: reference.sheetId,
      get protection(): SheetProtection | undefined {
        const pending = sheetProtections.get(reference.path)
        if (pending === 'remove') return undefined
        if (pending !== undefined) return pending
        return sheetBytes === undefined ? undefined : readSheetProtection(sheetBytes)
      },
      cells: () => readCells(),
      cell(cellReference: string): Cell | undefined {
        const wanted = canonicalReference(parseReference(cellReference))
        if (wanted === undefined) return undefined
        return findCell(wanted)
      },
      set(cellReference: string, value: CellInput, options?: SetOptions): void {
        // Normalised so `a1`, `$A$1` and `A1` are one edit, and so the file
        // never receives a reference spelled the way the caller typed it.
        const canonical = formatReference(parseWritableReference(cellReference))

        // Refused here rather than at save time. An edit that only fails once the
        // workbook is written takes the whole batch down with it, and until then
        // cell() reports a write that is never going to happen.
        if (sheetBytes === undefined) throw absent(canonical, 'written')

        checkWritable(canonical, value, date1904, at)
        checkStyleOptions(options, canonical)
        // sheetBytes is present, so indexed() is too; the guard is for the type.
        const index = indexed()
        if (index !== undefined) {
          const si = index.sharedFormulas.get(canonical)
          if (si !== undefined) throw sharedFormulaRefusal(canonical, si, at)
          const anchor = mergeAnchorFor(index, canonical)
          if (anchor !== undefined) throw mergeRefusal(canonical, anchor, at)
        }

        const current = styleAt(canonical)
        const applied = resolveStyle(current, value, options, canonical)

        const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
        pending.set(canonical, value)
        edits.set(reference.path, pending)

        overlay.set(canonical, predict(canonical, value, commitStyle(canonical, current, applied)))
      },
      format(cellReference: string, options: SetOptions): void {
        const canonical = formatReference(parseWritableReference(cellReference))
        if (sheetBytes === undefined) throw absent(canonical, 'formatted')
        checkStyleOptions(options, canonical)

        const current = styleAt(canonical)
        // No value, so a Date never triggers a format; only what is asked for.
        // Overwriting a shared-formula master or a merged cell is safe here: the
        // value and formula are kept, so neither refusal a write needs applies.
        const applied = resolveStyle(current, undefined, options, canonical)
        if (applied === undefined) return

        const existing = findCell(canonical)
        const resolved = commitStyle(canonical, current, applied)
        overlay.set(
          canonical,
          existing === undefined
            ? predict(canonical, null, resolved)
            : restyled(existing, resolved),
        )
      },
      protect(options: SheetProtection = {}): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so it cannot be protected`,
            at,
          )
        }
        checkProtection(options, at)
        sheetProtections.set(reference.path, options)
      },
      unprotect(): void {
        // A missing sheet part is left to toBytes to skip, as it has nothing to
        // unprotect; recording the intent is harmless.
        sheetProtections.set(reference.path, 'remove')
      },
      merge(range: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${range} cannot be merged`,
            { ...at, reference: range },
          )
        }
        const canonical = mergeRangeReference(range, at)
        const pending = sheetMerges.get(reference.path) ?? []
        pending.push(canonical)
        sheetMerges.set(reference.path, pending)
      },
      autoFilter(range: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${range} cannot be filtered`,
            { ...at, reference: range },
          )
        }
        sheetAutoFilters.set(reference.path, mergeRangeReference(range, at))
      },
      freeze(cell: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so it cannot be frozen`,
            { ...at, reference: cell },
          )
        }
        sheetFreezes.set(reference.path, formatReference(parseWritableReference(cell)))
      },
      setRowHeight(row: number, height: number): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so row ${row} cannot be sized`,
            at,
          )
        }
        if (!Number.isInteger(row) || row < 1) {
          throw new XlsxError('bad-reference', `Row ${row} is not a row number`, { ...at })
        }
        if (row > LAST_ROW) {
          throw new XlsxError('bad-reference', `Row ${row} is outside the sheet`, { ...at })
        }
        if (!Number.isFinite(height) || height < 0) {
          throw new XlsxError('unwritable-value', `Row height ${height} is not zero or more`, {
            ...at,
          })
        }
        const heights = sheetRowHeights.get(reference.path) ?? new Map<number, number>()
        heights.set(row, height)
        sheetRowHeights.set(reference.path, heights)
      },
      setColumnWidth(column: string, width: number): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so column ${column} cannot be sized`,
            { ...at, reference: column },
          )
        }
        const index = columnToIndex(column)
        if (index > LAST_COLUMN) {
          throw new XlsxError('bad-reference', `Column ${column} is outside the sheet`, {
            ...at,
            reference: column,
          })
        }
        if (!Number.isFinite(width) || width < 0) {
          throw new XlsxError('unwritable-value', `Column width ${width} is not zero or more`, {
            ...at,
            reference: column,
          })
        }
        const widths = sheetColumnWidths.get(reference.path) ?? new Map<number, number>()
        widths.set(index, width)
        sheetColumnWidths.set(reference.path, widths)
      },
      hideRow(row: number): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so row ${row} cannot be hidden`,
            at,
          )
        }
        if (!Number.isInteger(row) || row < 1 || row > LAST_ROW) {
          throw new XlsxError('bad-reference', `Row ${row} is not a row this sheet can hold`, {
            ...at,
          })
        }
        const hidden = sheetHiddenRows.get(reference.path) ?? new Set<number>()
        hidden.add(row)
        sheetHiddenRows.set(reference.path, hidden)
      },
      hideColumn(column: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so column ${column} cannot be hidden`,
            { ...at, reference: column },
          )
        }
        const index = columnToIndex(column)
        if (index > LAST_COLUMN) {
          throw new XlsxError('bad-reference', `Column ${column} is outside the sheet`, {
            ...at,
            reference: column,
          })
        }
        const hidden = sheetHiddenColumns.get(reference.path) ?? new Set<number>()
        hidden.add(index)
        sheetHiddenColumns.set(reference.path, hidden)
      },
      insertRows(before: number, count = 1): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so no rows can be inserted`,
            at,
          )
        }
        if (!Number.isInteger(before) || before < 1 || before > LAST_ROW) {
          throw new XlsxError('bad-reference', `Row ${before} is not a row this sheet can hold`, {
            ...at,
          })
        }
        checkCount(count, 'rows to insert')
        if (highestRow(sheetBytes) + count > LAST_ROW) {
          throw new XlsxError(
            'unwritable-value',
            `Inserting ${count} row(s) would push a row off ${reference.name}`,
            { ...at },
          )
        }
        refuseUnshiftable('its rows cannot be inserted into yet')
        lineOps.push({ path: reference.path, spec: lineSpec('row', before, count) })
      },
      insertColumns(before: string, count = 1): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so no columns can be inserted`,
            { ...at, reference: before },
          )
        }
        const atColumn = columnToIndex(before)
        if (atColumn > LAST_COLUMN) {
          throw new XlsxError('bad-reference', `Column ${before} is outside the sheet`, {
            ...at,
            reference: before,
          })
        }
        checkCount(count, 'columns to insert', before)
        if (highestColumn(sheetBytes) + count > LAST_COLUMN) {
          throw new XlsxError(
            'unwritable-value',
            `Inserting ${count} column(s) would push a column off ${reference.name}`,
            { ...at, reference: before },
          )
        }
        refuseUnshiftable('its columns cannot be inserted into yet', before)
        lineOps.push({ path: reference.path, spec: lineSpec('column', atColumn, count) })
      },
      deleteRows(from: number, count = 1): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so no rows can be deleted`,
            at,
          )
        }
        if (!Number.isInteger(from) || from < 1 || from > LAST_ROW) {
          throw new XlsxError('bad-reference', `Row ${from} is not a row this sheet can hold`, {
            ...at,
          })
        }
        checkCount(count, 'rows to delete')
        refuseUnshiftable('its rows cannot be deleted yet')
        const spec = lineSpec('row', from, -count)
        const damage = deletionDamage(sheetBytes, spec)
        if (damage !== undefined) {
          throw new XlsxError(
            'unwritable-value',
            `Deleting rows from ${reference.name} would destroy ${damage}`,
            { ...at },
          )
        }
        lineOps.push({ path: reference.path, spec })
      },
      deleteColumns(from: string, count = 1): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so no columns can be deleted`,
            { ...at, reference: from },
          )
        }
        const atColumn = columnToIndex(from)
        if (atColumn > LAST_COLUMN) {
          throw new XlsxError('bad-reference', `Column ${from} is outside the sheet`, {
            ...at,
            reference: from,
          })
        }
        checkCount(count, 'columns to delete', from)
        refuseUnshiftable('its columns cannot be deleted yet', from)
        const spec = lineSpec('column', atColumn, -count)
        const damage = deletionDamage(sheetBytes, spec)
        if (damage !== undefined) {
          throw new XlsxError(
            'unwritable-value',
            `Deleting columns from ${reference.name} would destroy ${damage}`,
            { ...at, reference: from },
          )
        }
        lineOps.push({ path: reference.path, spec })
      },
      link(cell: string, target: Hyperlink): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${cell} cannot be linked`,
            { ...at, reference: cell },
          )
        }
        const canonical = formatReference(parseWritableReference(cell))
        const place = 'url' in target ? target.url : target.location
        if (typeof place !== 'string' || place.length === 0) {
          throw new XlsxError('unwritable-value', `A hyperlink needs a non-empty url or location`, {
            ...at,
            reference: canonical,
          })
        }
        if (target.tooltip !== undefined && typeof target.tooltip !== 'string') {
          throw new XlsxError('unwritable-value', `A hyperlink tooltip must be a string`, {
            ...at,
            reference: canonical,
          })
        }
        const links = sheetHyperlinks.get(reference.path) ?? new Map<string, Hyperlink>()
        links.set(canonical, target)
        sheetHyperlinks.set(reference.path, links)
      },
    }
  }

  const sheets: Worksheet[] = part.sheets.map(makeWorksheet)

  const workbookDir = part.path.replace(/[^/]+$/, '')
  const usedSheetPaths = new Set(container.parts.keys())
  let maxSheetId = part.sheets.reduce((max, sheet) => Math.max(max, Number(sheet.sheetId) || 0), 0)
  const workbookRelsXml = partText(container, part.relationshipsPath) ?? ''
  let maxRelationshipId = 0
  for (const match of workbookRelsXml.matchAll(/Id="rId(\d+)"/g)) {
    maxRelationshipId = Math.max(maxRelationshipId, Number(match[1]))
  }

  const addWorksheet = (name: string): Worksheet => {
    checkSheetName(
      name,
      sheets.map((sheet) => sheet.name),
    )
    let n = 1
    while (usedSheetPaths.has(`${workbookDir}worksheets/sheet${n}.xml`)) n++
    const path = `${workbookDir}worksheets/sheet${n}.xml`
    usedSheetPaths.add(path)
    const reference: SheetRef = {
      name,
      path,
      sheetId: String(++maxSheetId),
      state: 'visible',
    }
    addedSheets.set(path, new TextEncoder().encode(EMPTY_SHEET_XML))
    addedRefs.push({
      reference,
      relationshipId: `rId${++maxRelationshipId}`,
      // The sheet path is built from the workbook's own folder, so the target is
      // just the tail past it, relative to where the workbook rels resolve from.
      target: path.slice(workbookDir.length),
    })
    const sheet = makeWorksheet(reference)
    sheets.push(sheet)
    return sheet
  }

  const toBytes = (): Uint8Array => {
    // A change carries new bytes for a part; null deletes it. Every part not
    // named here is passed through still compressed, never inflated or rebuilt.
    const changes = new Map<string, Uint8Array | null>()
    // A format() with no set() records a style override and no value edit; a
    // protect(), merge() or setRowHeight() records neither and still rewrites.
    if (
      patchInputs.every((map) => map.size === 0) &&
      sheetHyperlinks.size === 0 &&
      addedRefs.length === 0 &&
      renames.size === 0 &&
      removed.size === 0 &&
      pendingNames.size === 0 &&
      lineOps.length === 0
    ) {
      return container.write(changes)
    }

    const encoder = new TextEncoder()

    // Excel rebuilds the calculation chain, but a stale one makes it offer to
    // repair the file. Dropping it also touches the workbook rels and the content
    // types, which the added-sheet wiring below writes in the same pass.
    const hadCalcChain = container.parts.has(CALCULATION_CHAIN)
    if (hadCalcChain) changes.set(CALCULATION_CHAIN, null)

    // set() already resolved every style; only the serialising is left.
    if (workingStyles !== undefined && workingStyles !== stylesXml) {
      changes.set('xl/styles.xml', encoder.encode(workingStyles))
    }

    // Text goes into the shared string table when the file has one, so the same
    // words written into many cells are stored once.
    let indexes: ReadonlyMap<string, number> | undefined
    if (stringsXml !== undefined) {
      const written: string[] = []
      for (const pending of edits.values()) {
        for (const value of pending.values()) {
          if (typeof value === 'string') written.push(value)
        }
      }
      if (written.length > 0) {
        const appended = appendSharedStrings(stringsXml, written)
        changes.set('xl/sharedStrings.xml', encoder.encode(appended.xml))
        indexes = appended.indexes
      }
    }

    // Every sheet with a value edit, a format() restyle, a protect(), a merge()
    // or a setRowHeight() is rewritten once.
    for (const path of new Set([
      ...patchInputs.flatMap((map) => [...map.keys()]),
      ...addedSheets.keys(),
    ])) {
      if (removed.has(path)) continue
      const bytes = container.parts.get(path) ?? addedSheets.get(path)
      if (bytes === undefined) continue
      const pending = edits.get(path) ?? EMPTY_EDITS
      const at: SheetLocation = {
        sheet:
          part.sheets.find((s) => s.path === path)?.name ??
          addedRefs.find((added) => added.reference.path === path)?.reference.name,
        part: path,
      }
      changes.set(
        path,
        patchSheet(bytes, pending, date1904, indexes, styleOverrides.get(path), at, {
          protection: sheetProtections.get(path),
          merges: sheetMerges.get(path),
          rowHeights: sheetRowHeights.get(path),
          columnWidths: sheetColumnWidths.get(path),
          autoFilter: sheetAutoFilters.get(path),
          freeze: sheetFreezes.get(path),
          hiddenRows: sheetHiddenRows.get(path),
          hiddenColumns: sheetHiddenColumns.get(path),
        }),
      )
      // A cell written just past a table grows it, the way Excel would.
      for (const extension of extendTables(bytes, path, container, pending.keys())) {
        changes.set(extension.path, encoder.encode(extension.xml))
      }
    }

    // A moved reference leaves the cached results stale, so a shift asks for a
    // recalculation just as writing a formula does.
    const wroteFormula =
      lineOps.length > 0 ||
      [...edits.values()].some((pending) =>
        [...pending.values()].some(
          (value) => typeof value === 'object' && value !== null && !(value instanceof Date),
        ),
      )

    // A removed sheet the file had is dropped and unwired below; a removed added
    // one was already taken out of the added set, so it never reaches here.
    const removedExisting = [...removed].filter((path) => container.parts.has(path))
    for (const path of removedExisting) changes.set(path, null)
    const originalName = (path: string) => part.sheets.find((sheet) => sheet.path === path)?.name

    const decodePart = new TextDecoder()
    const sheetTextNow = (path: string): string | undefined => {
      const changed = changes.get(path)
      if (changed !== undefined && changed !== null) return decodePart.decode(changed)
      return partText(container, path)
    }

    // Composes a top-level part written once this pass — the workbook, its rels,
    // the content types — from its current text, and writes it back only when the
    // transform changed something, so an untouched part is never rewritten.
    const rewritePart = (path: string, transform: (xml: string) => string): void => {
      const xml = partText(container, path)
      if (xml === undefined) return
      const updated = transform(xml)
      if (updated !== xml) changes.set(path, encoder.encode(updated))
    }

    // Hyperlinks are written before the line-ops below, so an inserted or deleted
    // line shifts their `ref` along with every other reference. An external link
    // takes a fresh relationship id in the sheet's rels part; an internal one is
    // written inline and needs none.
    for (const [path, links] of sheetHyperlinks) {
      if (removed.has(path)) continue
      const sheetXml = sheetTextNow(path)
      if (sheetXml === undefined) continue
      const relationshipsPath = relationshipsPathFor(path)
      const existingRels = partText(container, relationshipsPath)
      let nextId = 0
      for (const match of (existingRels ?? '').matchAll(/Id="rId(\d+)"/g)) {
        nextId = Math.max(nextId, Number(match[1]))
      }
      const entries: HyperlinkEntry[] = []
      const externalRels: { id: string; url: string }[] = []
      for (const [reference, target] of links) {
        if ('url' in target) {
          const id = `rId${++nextId}`
          entries.push({ reference, relationshipId: id, tooltip: target.tooltip })
          externalRels.push({ id, url: target.url })
        } else {
          entries.push({ reference, location: target.location, tooltip: target.tooltip })
        }
      }
      changes.set(path, encoder.encode(withHyperlinks(sheetXml, entries)))
      if (externalRels.length > 0) {
        changes.set(
          relationshipsPath,
          encoder.encode(withHyperlinkRelationships(existingRels, externalRels)),
        )
      }
    }

    // Inserting or deleting a line moves references across the whole workbook. The
    // per-sheet patch above has already landed this session's edits in the old
    // grid, so each sheet's current text is shifted here: the edited sheet's own
    // rows and cells, and every other sheet's formulas that point into it. Defined
    // names, global to the workbook, shift into the set written below.
    let namesToWrite: ReadonlyMap<string, string> = pendingNames
    if (lineOps.length > 0) {
      const sheetPaths = [...part.sheets.map((sheet) => sheet.path), ...addedSheets.keys()]
      for (const path of sheetPaths) {
        if (removed.has(path)) continue
        let xml = sheetTextNow(path)
        if (xml === undefined) continue
        const before = xml
        for (const op of lineOps) {
          xml =
            op.path === path
              ? shiftSheet(xml, op.spec)
              : shiftForeignFormulas(xml, { ...op.spec, onCurrentSheet: false })
        }
        if (xml !== before) changes.set(path, encoder.encode(xml))
      }

      let names = new Map([...fileNames, ...pendingNames])
      for (const op of lineOps) {
        names = new Map(
          [...names].map(([name, refersTo]) => [
            name,
            shiftFormula(refersTo, { ...op.spec, onCurrentSheet: false }),
          ]),
        )
      }
      const written = new Map(pendingNames)
      for (const [name, refersTo] of names) {
        if (pendingNames.has(name) || refersTo !== fileNames.get(name)) written.set(name, refersTo)
      }
      namesToWrite = written
    }

    // The workbook, its relationships and the content types each take a handful of
    // edits — the calculation chain, a recalculation flag, the sheets added,
    // renamed or removed — so each is composed from its current text and written
    // once. An added sheet is renamed by the name its wiring is written with; a
    // sheet the file already had by rewriting its existing <sheet>.
    const renamedAdded = addedRefs.map((added) => ({
      ...added,
      reference: {
        ...added.reference,
        name: renames.get(added.reference.path) ?? added.reference.name,
      },
    }))
    rewritePart(part.path, (workbookXml) => {
      let updated = wroteFormula ? withRecalculation(workbookXml) : workbookXml
      for (const [path, name] of renames) {
        const original = originalName(path)
        if (original !== undefined && !removed.has(path)) {
          updated = withSheetRenamed(updated, original, name)
        }
      }
      updated = withSheetsAdded(updated, renamedAdded)
      for (const path of removedExisting) {
        const original = originalName(path)
        if (original !== undefined) updated = withSheetRemoved(updated, original)
      }
      return withDefinedNames(updated, namesToWrite)
    })

    rewritePart(part.relationshipsPath, (relationshipsXml) => {
      let updated = withSheetRelationships(
        hadCalcChain
          ? withoutRelationship(relationshipsXml, part.path, CALCULATION_CHAIN)
          : relationshipsXml,
        addedRefs,
      )
      for (const path of removedExisting) {
        updated = withoutRelationship(updated, part.path, path)
      }
      return updated
    })

    rewritePart(CONTENT_TYPES, (contentTypesXml) => {
      let updated = withSheetContentTypes(
        hadCalcChain ? withoutOverride(contentTypesXml, CALCULATION_CHAIN) : contentTypesXml,
        addedRefs,
      )
      for (const path of removedExisting) updated = withoutOverride(updated, path)
      return updated
    })

    return container.write(changes)
  }

  return {
    sheets,
    sheet: (name: string) => sheets.find((candidate) => candidate.name === name),
    addSheet: addWorksheet,
    get definedNames(): ReadonlyMap<string, string> {
      return new Map([...fileNames, ...pendingNames])
    },
    defineName(name: string, refersTo: string): void {
      checkDefinedName(name, refersTo)
      pendingNames.set(name, refersTo)
    },
    epoch: date1904 ? 1904 : 1900,
    toBytes,
  }
}
