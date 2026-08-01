import {
  type AddedSheet,
  EMPTY_SHEET_XML,
  checkSheetName,
  withSheetContentTypes,
  withSheetRelationships,
  withSheetRemoved,
  withSheetRenamed,
  withSheetState,
  withSheetsAdded,
} from './add-sheet.js'
import { blankWorkbookBytes } from './blank.js'
import { type ContainerDraft, createContainerDraft, withRelationship } from './container-draft.js'
import { withContentTypeOverride } from './content-types.js'
import { applyLineShifts } from './apply-line-shifts.js'
import { createConditionalFormatStore, planConditionalFormat } from './conditional-format.js'
import { formatCsv, parseCsv } from './csv.js'
import { createValidationStore } from './data-validation.js'
import { type PendingImage, contributeImages, imageType } from './images.js'
import { COMMENTS_RELATIONSHIP, contributeComments, readComments } from './comments.js'
import {
  PRINT_AREA_NAME,
  type SheetScopedName,
  checkDefinedName,
  isBuiltInName,
  printAreaRanges,
  printAreaRefersTo,
  readDefinedNames,
  readSheetScopedNames,
  withDefinedNames,
} from './defined-names.js'
import {
  type DocumentProperties,
  readCoreProperties,
  writeCoreProperties,
} from './document-properties.js'
import { readRelationships, resolveTarget } from './relationships.js'
import { type Hyperlink, contributeHyperlinks, readSheetHyperlinks } from './hyperlinks.js'
import { type Container, decodeXmlPart } from './container.js'
import { XlsxError } from './errors.js'
import { LAST_SERIAL, dateToSerial, serialToDate } from './date.js'
import type { CellInput, SheetLocation } from './cell-input.js'
import {
  type ConditionalFormatSpec,
  type DataValidationSpec,
  checkProtection,
  checkWritable,
  mergeAnchorFor,
  mergeRangeReference,
  mergeRefusal,
  withoutConditionalFormatting,
  withoutDataValidations,
  withoutMergeCells,
  patchSheet,
  indexSheet,
  readColumnGroupLevels,
  readColumnWidths,
  readConditionalFormats,
  readDataValidations,
  readHiddenColumns,
  readHiddenRows,
  readRowGroupLevels,
  readRowHeights,
  readSheetProtection,
  readSheetView,
  sharedFormulaRefusal,
  type SheetIndex,
  type SheetViewState,
  type SheetProtection,
} from './patch.js'
import {
  LAST_COLUMN,
  LAST_ROW,
  canonicalReference,
  columnToIndex,
  formatReference,
  parseReference,
  parseWritableReference,
} from './reference.js'
import { type SharedMasters, toCell } from './cell-read.js'
import { flattenRuns, richTextOf } from './rich-text.js'
import { readSheet } from './sheet.js'
import { appendSharedStrings, readRichSharedStrings, readSharedStrings } from './shared-strings.js'
import {
  type TableSpec,
  contributeTables,
  extendTables,
  planTable,
  readTables,
  tableColumnDamage,
  tableRowDamage,
} from './tables.js'
import type { HeaderFooter, PageBreaks, PageMargins, PageSetup } from './page.js'
import { createPageStore } from './page-session.js'
import {
  type CellFormatting,
  type Color,
  checkStyleOptions,
  ensureDxf,
  normalizeColor,
} from './styles-writer.js'
import { type StylesSession, createStylesSession } from './styles-session.js'
import type { ShiftSpec } from './shift.js'
import { paletteByIndex, readThemeColors, resolveColor } from './theme.js'
import { type Styles, isDateFormat, numberFormatOf } from './styles.js'
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
import type {
  Cell,
  CellValue,
  ConditionalFormat,
  CsvReadOptions,
  DataValidation,
  SetOptions,
  Workbook,
  Worksheet,
} from './public-types.js'

const EMPTY_STYLES: Styles = { numberFormats: new Map(), cellFormats: [] }
const EMPTY_EDITS: ReadonlyMap<string, CellInput> = new Map()

const CALCULATION_CHAIN = 'xl/calcChain.xml'
const CONTENT_TYPES = '[Content_Types].xml'
const ROOT_RELATIONSHIPS = '_rels/.rels'
const CORE_PROPERTIES_PART = 'docProps/core.xml'
const CORE_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const CORE_PROPERTIES_CONTENT_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'

// Pinned parts whose positions this library now moves with an insert or delete,
// so their presence no longer blocks one at the gate. A table still refuses a
// column edit that would resize it — that check is tableColumnDamage, once the
// edit's position is known. A drawing, pivot table or chart is not here yet.
const SHIFTABLE_PARTS: ReadonlySet<string> = new Set([
  'a table',
  'a comment',
  'a legacy drawing',
  'a drawing',
  'a pivot table',
])

function partText(container: Container, path: string): string | undefined {
  const bytes = container.parts.get(path)
  if (bytes === undefined) return undefined
  return decodeXmlPart(bytes, path)
}

/** Canonicalises a cell or range into the `sqref` an element takes. */
function sqrefOf(range: string, at: SheetLocation): string {
  if (range.includes(':')) return mergeRangeReference(range, at)
  return formatReference(parseWritableReference(range))
}

// A worksheet's <legacyDrawing> must sit after the drawing elements and before
// these, which the schema orders after it. Inserting before the earliest present
// keeps the order valid; a worksheet-level <extLst> is always the last child, so
// it is handled separately from any <extLst> nested in an earlier element.
function checkOutlineLevel(level: number, at: SheetLocation): void {
  if (!Number.isInteger(level) || level < 1 || level > 7) {
    throw new XlsxError('unwritable-value', `Outline level ${level} is not between 1 and 7`, {
      ...at,
    })
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

const CSV_NUMBER = /^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/

/** One cell's value as a CSV field: a number as itself, a date in ISO form, a
 * boolean as TRUE/FALSE, and an empty cell or an uncalculated formula as blank. */
function csvField(value: CellValue): string {
  switch (value.kind) {
    case 'number':
      return String(value.value)
    case 'text':
    case 'error':
      return value.value
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE'
    case 'date':
      return value.value.toISOString()
    case 'empty':
      return ''
  }
}

/**
 * Builds a workbook from delimited text: one sheet whose cells hold the parsed
 * fields from `A1`. Fields are text unless `parseNumbers` turns a numeric one into
 * a number. The result is a normal created workbook — fill it further and write it
 * with `toBytes()`.
 */
export function createWorkbookFromCsv(text: string, options: CsvReadOptions = {}): Workbook {
  const rows = parseCsv(text, options.delimiter)
  const workbook = createWorkbook(options.sheetName)
  const data: CellInput[][] = rows.map((row) =>
    row.map((field) =>
      options.parseNumbers === true && CSV_NUMBER.test(field) && Number.isFinite(Number(field))
        ? Number(field)
        : field,
    ),
  )
  if (data.length > 0) workbook.sheets[0]?.setValues('A1', data)
  return workbook
}

export function readWorkbook(bytes: Uint8Array): Workbook {
  const part = readWorkbookPart(bytes)
  const { container, date1904 } = part

  const stylesXml = partText(container, 'xl/styles.xml')

  const stringsXml = partText(container, 'xl/sharedStrings.xml')
  const sharedStrings = stringsXml === undefined ? [] : readSharedStrings(stringsXml)
  const richStrings = stringsXml === undefined ? [] : readRichSharedStrings(stringsXml)

  const themeXml = partText(container, 'xl/theme/theme1.xml')
  const themePalette = themeXml === undefined ? [] : paletteByIndex(readThemeColors(themeXml))

  const edits = new Map<string, Map<string, CellInput>>()

  // Which cell format each edited cell lands on is decided by set(), not by
  // toBytes(). Choosing a style index is what makes a number a date, so it is a
  // decision about what a cell MEANS; leaving it to write time gave the read
  // path its own copy of the decision, and the two drifted.
  const styleOverrides = new Map<string, Map<string, number>>()
  const sheetProtections = new Map<string, SheetProtection | 'remove'>()
  const sheetMerges = new Map<string, string[]>()
  const sheetUnmerges = new Map<string, Set<string>>()
  const sheetRowHeights = new Map<string, Map<number, number>>()
  const sheetColumnWidths = new Map<string, Map<number, number>>()
  const sheetAutoFilters = new Map<string, string>()
  const sheetFreezes = new Map<string, string>()
  const sheetStates = new Map<string, SheetState>()
  const sheetHiddenRows = new Map<string, Set<number>>()
  const sheetHiddenColumns = new Map<string, Set<number>>()
  const sheetTabColors = new Map<string, string>()
  const sheetGridlines = new Map<string, boolean>()
  const sheetHeadings = new Map<string, boolean>()
  const sheetZoom = new Map<string, number>()
  const sheetRowGroups = new Map<string, Map<number, number>>()
  const sheetColGroups = new Map<string, Map<number, number>>()
  const validations = createValidationStore(date1904)
  const conditionalFormats = createConditionalFormatStore()
  // Comments to add, per sheet, only for sheets that had none — an existing
  // comments part is refused at the call rather than rebuilt.
  const sheetComments = new Map<string, Map<string, string>>()
  const sheetRemovedComments = new Map<string, Set<string>>()
  const sheetTables = new Map<string, TableSpec[]>()
  // Images to embed, per sheet; toBytes builds or extends the sheet's drawing.
  const sheetImages = new Map<string, PendingImage[]>()
  const page = createPageStore()
  // The per-sheet maps patchSheet applies in one rewrite. Both the "anything
  // pending?" check and the set of sheets to rewrite read this list, so a new
  // kind of sheet edit is registered in one place rather than two enumerations
  // that have to be kept in step.
  const patchInputs: ReadonlyArray<ReadonlyMap<string, unknown>> = [
    edits,
    styleOverrides,
    sheetProtections,
    sheetMerges,
    sheetUnmerges,
    sheetRowHeights,
    sheetColumnWidths,
    sheetAutoFilters,
    sheetFreezes,
    sheetHiddenRows,
    sheetHiddenColumns,
    sheetTabColors,
    sheetGridlines,
    sheetHeadings,
    sheetZoom,
    sheetRowGroups,
    sheetColGroups,
    validations.pending,
    conditionalFormats.pending,
  ]
  const fileNames = readDefinedNames(partText(container, part.path) ?? '')
  const pendingNames = new Map<string, string>()
  const removedNames = new Set<string>()
  // Sheet-scoped names, keyed by the sheet's package path (stable across a rename
  // or a reorder). The file's own are read once; toBytes maps a path to its final
  // localSheetId. A built-in `_xlnm.*` name is reserved for its own accessor.
  const fileScopedNames = readSheetScopedNames(partText(container, part.path) ?? '')
  const sheetPendingNames = new Map<string, Map<string, string>>()
  const sheetRemovedNames = new Map<string, Set<string>>()
  let pendingProperties: DocumentProperties = {}
  const sheetHyperlinks = new Map<string, Map<string, Hyperlink>>()
  const sheetUnlinks = new Map<string, Set<string>>()
  // Row and column inserts and deletes, in call order. Each names the sheet it
  // was called on; toBytes applies them after the per-sheet patch so an edit made
  // this session lands in the old grid and then moves with the shift.
  const lineOps: { readonly path: string; readonly spec: ShiftSpec }[] = []
  // Style edits amend an in-memory model instead of re-parsing and re-splicing a
  // growing string on every set(); the string is produced once at toBytes. A file
  // with no style table has no session, and any format() on it is refused.
  const session: StylesSession | undefined =
    stylesXml === undefined ? undefined : createStylesSession(stylesXml)

  const stylesNow = (): Styles => session?.styles() ?? EMPTY_STYLES

  const formattingFor = (styleIndex: number | undefined): CellFormatting =>
    styleIndex === undefined || session === undefined ? {} : session.formattingOf(styleIndex)

  // Conditional-format highlights live in dxfs, a table the session leaves alone.
  // The path is low volume, so it keeps threading a string; the ordered colours
  // let toBytes fold the same dxfs onto the serialized styles.
  let dxfStyles = stylesXml
  const dxfColors: string[] = []

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
  // Whole parts a copySheet brings into the package — the duplicated dependent
  // parts and the copy's own relationships — written verbatim by toBytes, plus the
  // content-type Overrides any of them need.
  const addedParts = new Map<string, Uint8Array>()
  const addedOverrides: Array<{ readonly path: string; readonly contentType: string }> = []

  const makeWorksheet = (reference: SheetRef): Worksheet => {
    const sheetBytes = container.parts.get(reference.path) ?? addedSheets.get(reference.path)
    const at: SheetLocation = { sheet: reference.name, part: reference.path }
    // Where the sheet sits in the file's order, which is the localSheetId its
    // file-scoped names carry. -1 for a sheet added this session, which has none.
    const fileSheetIndex = part.sheets.findIndex((sheet) => sheet.path === reference.path)

    // The sheet points at its comments part through a relationship. Found once, it
    // gives what each cell reports and which cells are already taken, so a second
    // comment on one of them is refused rather than rebuilding its rich text.
    const existingCommentsPath = (() => {
      const relsXml = partText(container, relationshipsPathFor(reference.path))
      if (relsXml === undefined) return undefined
      for (const relationship of readRelationships(relsXml, reference.path).values()) {
        if (relationship.type === COMMENTS_RELATIONSHIP && !relationship.external) {
          return resolveTarget(reference.path, relationship.target)
        }
      }
      return undefined
    })()
    const commentsRead =
      existingCommentsPath === undefined
        ? new Map<string, string>()
        : readComments(partText(container, existingCommentsPath) ?? '')

    // Read on first use from the same bytes the cells stream from, and memoised;
    // unlike comments this parses the whole sheet, so a workbook whose cells are
    // never read does not pay for it.
    let viewCache: SheetViewState | undefined
    const viewState = (): SheetViewState => {
      if (viewCache === undefined) {
        viewCache =
          sheetBytes === undefined ? { gridlines: true, headings: true } : readSheetView(sheetBytes)
      }
      return viewCache
    }
    let rowHeightsCache: ReadonlyMap<number, number> | undefined
    let columnWidthsCache: readonly { min: number; max: number; width: number }[] | undefined
    let hiddenRowsCache: ReadonlySet<number> | undefined
    let hiddenColumnsCache: readonly { min: number; max: number }[] | undefined
    let rowGroupsCache: ReadonlyMap<number, number> | undefined
    let columnGroupsCache: readonly { min: number; max: number; level: number }[] | undefined
    let validationsCache: readonly DataValidationSpec[] | undefined
    let conditionalFormatsCache: readonly ConditionalFormatSpec[] | undefined
    let hyperlinksCache: ReadonlyMap<string, Hyperlink> | undefined
    const hyperlinksFor = (bytes: Uint8Array): ReadonlyMap<string, Hyperlink> => {
      if (hyperlinksCache === undefined) {
        const relationshipsPath = relationshipsPathFor(reference.path)
        hyperlinksCache = readSheetHyperlinks(
          decodeXmlPart(bytes, reference.path),
          partText(container, relationshipsPath),
          relationshipsPath,
        )
      }
      return hyperlinksCache
    }

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

      // A comment or a link can sit on a cell the sheet never gave a <c>, so the
      // refs carrying one are tracked and any left after the stream are surfaced
      // as empty cells below. Only paid for when the sheet has some.
      const links = hyperlinksFor(bytes)
      const annotate = (cell: Cell, reference: string): Cell => {
        const comment = commentsRead.get(reference)
        const hyperlink = links.get(reference)
        if (comment === undefined && hyperlink === undefined) return cell
        return {
          ...cell,
          ...(comment === undefined ? {} : { comment }),
          ...(hyperlink === undefined ? {} : { hyperlink }),
        }
      }
      const unplaced =
        commentsRead.size + links.size > 0
          ? new Set<string>([...commentsRead.keys(), ...links.keys()])
          : undefined
      for (const raw of readSheet(bytes, sharedStrings, at, richStrings)) {
        if (raw.ownsSharedRange === true && raw.sharedIndex !== undefined) {
          masters.set(raw.sharedIndex, canonicalReference(raw.address) ?? raw.reference)
        }
        const cell = toCell(raw, stylesNow(), formattingFor(raw.styleIndex), date1904, masters)
        unplaced?.delete(cell.reference)
        yield annotate(cell, cell.reference)
      }
      for (const reference of unplaced ?? []) {
        const address = parseReference(reference)
        const base: Cell = {
          address,
          reference: canonicalReference(address) ?? reference,
          value: { kind: 'empty' },
        }
        yield annotate(base, reference)
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
      if ('runs' in value) {
        // Written inline, so it reads back the same: the flattened text as the
        // value, the runs alongside when they are meaningfully rich.
        const rich = richTextOf(value.runs)
        return toCell(
          {
            ...raw,
            value: { kind: 'text', value: flattenRuns(value.runs) },
            ...(rich === undefined ? {} : { richText: rich }),
          },
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

    // Writes a table's header text into a cell, the same way `set` records a plain
    // string, so the header row Excel requires matches the columns declared.
    const writeHeaderCell = (canonical: string, name: string): void => {
      checkWritable(canonical, name, date1904, at)
      const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
      pending.set(canonical, name)
      edits.set(reference.path, pending)
      overlay.set(canonical, predict(canonical, name, styleAt(canonical)))
    }

    const headerTextOf = (cell: Cell | undefined): string | undefined => {
      if (cell === undefined) return undefined
      if (cell.value.kind === 'text') return cell.value.value === '' ? undefined : cell.value.value
      if (cell.value.kind === 'number') return String(cell.value.value)
      return undefined
    }

    // Table names are unique across the workbook, so a fresh one and a collision
    // check both need every name already taken — the file's and this session's.
    const takenTableNames = (): Set<string> => {
      const names = new Set<string>()
      for (const sheet of part.sheets) {
        const bytes = container.parts.get(sheet.path)
        if (bytes !== undefined) {
          for (const table of readTables(bytes, sheet.path, container))
            names.add(table.name.toLowerCase())
        }
      }
      for (const specs of sheetTables.values()) {
        for (const spec of specs) names.add(spec.name.toLowerCase())
      }
      return names
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
    ): number | undefined => {
      const location = { ...at, reference: canonical }
      if (session === undefined) {
        if (Object.values(options ?? {}).some((asked) => asked !== undefined)) {
          throw new XlsxError(
            'missing-part',
            `Cannot format ${canonical}: the package has no style table`,
            { ...location, part: 'xl/styles.xml' },
          )
        }
        return undefined
      }
      // One transaction, so a step that throws part-way rolls back everything the
      // earlier steps added — the file stays as if nothing was asked.
      return session.transaction(() => {
        let base = current
        let applied = false
        const step = (index: number) => {
          base = index
          applied = true
        }
        // An asked-for format wins; a Date only gets one because without one it
        // displays as the serial number it is stored as.
        if (options?.numberFormat !== undefined)
          step(session.numberFormat(base, options.numberFormat, location))
        else if (value instanceof Date) step(session.dateStyle(base))
        if (options?.font !== undefined) step(session.font(base, options.font, location))
        if (options?.fill !== undefined) step(session.fill(base, options.fill, location))
        if (options?.border !== undefined) step(session.border(base, options.border, location))
        if (options?.alignment !== undefined)
          step(session.alignment(base, options.alignment, location))
        if (options?.protection !== undefined) step(session.protection(base, options.protection))
        return applied ? base : undefined
      })
    }

    const commitStyle = (
      canonical: string,
      current: number | undefined,
      applied: number | undefined,
    ): number | undefined => {
      if (applied === undefined) return current
      if (applied !== current) {
        const overrides = styleOverrides.get(reference.path) ?? new Map<string, number>()
        overrides.set(canonical, applied)
        styleOverrides.set(reference.path, overrides)
      }
      return applied
    }

    const absent = (canonical: string, verb: string): XlsxError =>
      new XlsxError(
        'missing-part',
        `Sheet ${reference.name} is not in the package, so ${canonical} cannot be ${verb}`,
        { ...at, reference: canonical },
      )

    // A part with pinned positions is preserved untouched, so an insert or delete
    // that would move the cells under it is refused until it too can be adjusted.
    const refuseUnshiftable = (
      action: string,
      where: string | undefined,
      allow?: ReadonlySet<string>,
    ): void => {
      const relationships = partText(container, relationshipsPathFor(reference.path))
      if (relationships === undefined) return
      // A drawing is shiftable only when it is readable and holds no chart, so it
      // is allowed at the gate and then checked here; the rest refuse by presence.
      const owns = unshiftablePart(relationships, allow) ?? drawingRefusal(relationships)
      if (owns === undefined) return
      throw new XlsxError(
        'unsupported-edit',
        `Sheet ${reference.name} carries ${owns}, so ${action}`,
        { ...at, ...(where === undefined ? {} : { reference: where }) },
      )
    }

    // A drawing's objects move by their cell anchor — a picture, shape, diagram or
    // OLE embed by position alone, a chart by its anchor and its series formulas —
    // since a chart is the only drawing object that references worksheet cells. So
    // the one drawing an edit refuses is one whose part cannot be read, and whose
    // anchors therefore cannot be moved to match the cells under them.
    const drawingRefusal = (relationshipsXml: string): string | undefined => {
      for (const relationship of readRelationships(relationshipsXml, reference.path).values()) {
        if (relationship.external || !relationship.type.endsWith('relationships/drawing')) continue
        if (partText(container, resolveTarget(reference.path, relationship.target)) === undefined)
          return 'a drawing'
      }
      return undefined
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

    // The one write path `set` and `setValues` share, so a block write is refused
    // and styled a cell at a time exactly as a single write is.
    const writeCell = (cellReference: string, value: CellInput, options?: SetOptions): void => {
      // Normalised so `a1`, `$A$1` and `A1` are one edit, and so the file never
      // receives a reference spelled the way the caller typed it.
      const canonical = formatReference(parseWritableReference(cellReference))

      // Refused here rather than at save time. An edit that only fails once the
      // workbook is written takes the whole batch down with it, and until then
      // cell() reports a write that is never going to happen.
      if (sheetBytes === undefined) throw absent(canonical, 'written')

      checkWritable(canonical, value, date1904, at)
      checkStyleOptions(options, canonical, at)
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
      get state(): SheetState {
        return sheetStates.get(reference.path) ?? reference.state
      },
      setState(state: SheetState): void {
        if (state !== 'visible') {
          const currentlyVisible =
            (sheetStates.get(reference.path) ?? reference.state) === 'visible'
          const visible = sheets.filter((candidate) => candidate.state === 'visible').length
          if (currentlyVisible && visible <= 1) {
            throw new XlsxError(
              'unsupported-edit',
              `Sheet ${reference.name} is the only visible sheet, so it cannot be hidden`,
              { ...at },
            )
          }
        }
        sheetStates.set(reference.path, state)
      },
      sheetId: reference.sheetId,
      get definedNames(): ReadonlyMap<string, string> {
        const result = new Map<string, string>()
        if (fileSheetIndex >= 0) {
          for (const entry of fileScopedNames) {
            if (entry.localSheetId === fileSheetIndex && !isBuiltInName(entry.name)) {
              result.set(entry.name, entry.refersTo)
            }
          }
        }
        for (const [name, refersTo] of sheetPendingNames.get(reference.path) ?? []) {
          if (!isBuiltInName(name)) result.set(name, refersTo)
        }
        for (const name of sheetRemovedNames.get(reference.path) ?? []) result.delete(name)
        return result
      },
      defineName(name: string, refersTo: string): void {
        checkDefinedName(name, refersTo)
        if (isBuiltInName(name)) {
          throw new XlsxError(
            'unwritable-value',
            `"${name}" is a built-in name Excel reserves; set the print area with its own method`,
            { ...at },
          )
        }
        sheetRemovedNames.get(reference.path)?.delete(name)
        const pending = sheetPendingNames.get(reference.path) ?? new Map<string, string>()
        pending.set(name, refersTo)
        sheetPendingNames.set(reference.path, pending)
      },
      removeDefinedName(name: string): void {
        sheetPendingNames.get(reference.path)?.delete(name)
        const removedSet = sheetRemovedNames.get(reference.path) ?? new Set<string>()
        removedSet.add(name)
        sheetRemovedNames.set(reference.path, removedSet)
      },
      get printArea(): string | undefined {
        const pending = sheetPendingNames.get(reference.path)?.get(PRINT_AREA_NAME)
        if (pending !== undefined) return printAreaRanges(pending)
        if (sheetRemovedNames.get(reference.path)?.has(PRINT_AREA_NAME)) return undefined
        if (fileSheetIndex >= 0) {
          for (const entry of fileScopedNames) {
            if (entry.localSheetId === fileSheetIndex && entry.name === PRINT_AREA_NAME) {
              return printAreaRanges(entry.refersTo)
            }
          }
        }
        return undefined
      },
      setPrintArea(range: string): void {
        const canonical = mergeRangeReference(range, at)
        const current = renames.get(reference.path) ?? reference.name
        sheetRemovedNames.get(reference.path)?.delete(PRINT_AREA_NAME)
        const pending = sheetPendingNames.get(reference.path) ?? new Map<string, string>()
        pending.set(PRINT_AREA_NAME, printAreaRefersTo(current, canonical))
        sheetPendingNames.set(reference.path, pending)
      },
      clearPrintArea(): void {
        sheetPendingNames.get(reference.path)?.delete(PRINT_AREA_NAME)
        const removedSet = sheetRemovedNames.get(reference.path) ?? new Set<string>()
        removedSet.add(PRINT_AREA_NAME)
        sheetRemovedNames.set(reference.path, removedSet)
      },
      get protection(): SheetProtection | undefined {
        const pending = sheetProtections.get(reference.path)
        if (pending === 'remove') return undefined
        if (pending !== undefined) return pending
        return sheetBytes === undefined ? undefined : readSheetProtection(sheetBytes)
      },
      get mergedRanges(): readonly string[] {
        const fromFile = (indexed()?.merges ?? []).map(
          (merge) =>
            `${merge.anchor}:${formatReference({ row: merge.maxRow, column: merge.maxColumn })}`,
        )
        const unmerged = sheetUnmerges.get(reference.path)
        const all = [...new Set([...fromFile, ...(sheetMerges.get(reference.path) ?? [])])]
        return unmerged === undefined ? all : all.filter((range) => !unmerged.has(range))
      },
      get gridlinesVisible(): boolean {
        return sheetGridlines.get(reference.path) ?? viewState().gridlines
      },
      get headingsVisible(): boolean {
        return sheetHeadings.get(reference.path) ?? viewState().headings
      },
      get zoomPercent(): number | undefined {
        return sheetZoom.get(reference.path) ?? viewState().zoom
      },
      get frozenAt(): string | undefined {
        return sheetFreezes.get(reference.path) ?? viewState().frozen
      },
      get tabColorHex(): string | undefined {
        return sheetTabColors.get(reference.path) ?? viewState().tabColor
      },
      get autoFilterRange(): string | undefined {
        return sheetAutoFilters.get(reference.path) ?? viewState().autoFilter
      },
      get validations(): readonly { readonly range: string; readonly rule: DataValidation }[] {
        if (validationsCache === undefined) {
          validationsCache = sheetBytes === undefined ? [] : readDataValidations(sheetBytes)
        }
        return validations.applied(reference.path, validationsCache)
      },
      get conditionalFormats(): readonly {
        readonly range: string
        readonly rule: ConditionalFormat
      }[] {
        if (conditionalFormatsCache === undefined) {
          conditionalFormatsCache =
            sheetBytes === undefined ? [] : readConditionalFormats(sheetBytes)
        }
        return conditionalFormats.applied(reference.path, conditionalFormatsCache, dxfStyles)
      },
      rowHeight(row: number): number | undefined {
        const pending = sheetRowHeights.get(reference.path)?.get(row)
        if (pending !== undefined) return pending
        if (sheetBytes === undefined) return undefined
        rowHeightsCache ??= readRowHeights(sheetBytes)
        return rowHeightsCache.get(row)
      },
      columnWidth(column: string): number | undefined {
        const index = columnToIndex(column)
        const pending = sheetColumnWidths.get(reference.path)?.get(index)
        if (pending !== undefined) return pending
        if (sheetBytes === undefined) return undefined
        columnWidthsCache ??= readColumnWidths(sheetBytes)
        return columnWidthsCache.find((range) => index >= range.min && index <= range.max)?.width
      },
      isRowHidden(row: number): boolean {
        if (sheetHiddenRows.get(reference.path)?.has(row) === true) return true
        if (sheetBytes === undefined) return false
        hiddenRowsCache ??= readHiddenRows(sheetBytes)
        return hiddenRowsCache.has(row)
      },
      isColumnHidden(column: string): boolean {
        const index = columnToIndex(column)
        if (sheetHiddenColumns.get(reference.path)?.has(index) === true) return true
        if (sheetBytes === undefined) return false
        hiddenColumnsCache ??= readHiddenColumns(sheetBytes)
        return hiddenColumnsCache.some((range) => index >= range.min && index <= range.max)
      },
      rowGroupLevel(row: number): number {
        const pending = sheetRowGroups.get(reference.path)?.get(row)
        if (pending !== undefined) return pending
        if (sheetBytes === undefined) return 0
        rowGroupsCache ??= readRowGroupLevels(sheetBytes)
        return rowGroupsCache.get(row) ?? 0
      },
      columnGroupLevel(column: string): number {
        const index = columnToIndex(column)
        const pending = sheetColGroups.get(reference.path)?.get(index)
        if (pending !== undefined) return pending
        if (sheetBytes === undefined) return 0
        columnGroupsCache ??= readColumnGroupLevels(sheetBytes)
        return (
          columnGroupsCache.find((range) => index >= range.min && index <= range.max)?.level ?? 0
        )
      },
      cells: () => readCells(),
      cell(cellReference: string): Cell | undefined {
        const wanted = canonicalReference(parseReference(cellReference))
        if (wanted === undefined) return undefined
        return findCell(wanted)
      },
      set: writeCell,
      setValues(
        topLeft: string,
        data: ReadonlyArray<ReadonlyArray<CellInput>>,
        options?: SetOptions,
      ): void {
        const origin = parseWritableReference(topLeft)
        data.forEach((row, rowOffset) => {
          row.forEach((value, columnOffset) => {
            const ref = formatReference({
              row: origin.row + rowOffset,
              column: origin.column + columnOffset,
            })
            writeCell(ref, value, options)
          })
        })
      },
      getValues(range: string): CellValue[][] {
        const colon = range.indexOf(':')
        const first = parseReference(colon === -1 ? range : range.slice(0, colon))
        const second = colon === -1 ? first : parseReference(range.slice(colon + 1))
        const minRow = Math.min(first.row, second.row)
        const maxRow = Math.max(first.row, second.row)
        const minColumn = Math.min(first.column, second.column)
        const maxColumn = Math.max(first.column, second.column)
        const rows: CellValue[][] = []
        for (let row = minRow; row <= maxRow; row++) {
          const values: CellValue[] = []
          for (let column = minColumn; column <= maxColumn; column++) {
            const found = findCell(formatReference({ row, column }))
            values.push(found?.value ?? { kind: 'empty' })
          }
          rows.push(values)
        }
        return rows
      },
      toCsv(options?: { readonly delimiter?: string }): string {
        let maxRow = 0
        let maxColumn = 0
        const byPosition = new Map<string, CellValue>()
        for (const found of readCells()) {
          if (found.value.kind === 'empty') continue
          maxRow = Math.max(maxRow, found.address.row)
          maxColumn = Math.max(maxColumn, found.address.column)
          byPosition.set(`${found.address.row},${found.address.column}`, found.value)
        }
        const rows: string[][] = []
        for (let row = 1; row <= maxRow; row++) {
          const fields: string[] = []
          for (let column = 1; column <= maxColumn; column++) {
            const value = byPosition.get(`${row},${column}`)
            fields.push(value === undefined ? '' : csvField(value))
          }
          rows.push(fields)
        }
        return formatCsv(rows, options?.delimiter)
      },
      format(cellReference: string, options: SetOptions): void {
        const canonical = formatReference(parseWritableReference(cellReference))
        if (sheetBytes === undefined) throw absent(canonical, 'formatted')
        checkStyleOptions(options, canonical, at)

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
        sheetUnmerges.get(reference.path)?.delete(canonical)
        const pending = sheetMerges.get(reference.path) ?? []
        pending.push(canonical)
        sheetMerges.set(reference.path, pending)
      },
      unmerge(range: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${range} cannot be unmerged`,
            { ...at, reference: range },
          )
        }
        const canonical = mergeRangeReference(range, at)
        const pending = sheetMerges.get(reference.path)
        if (pending !== undefined) {
          sheetMerges.set(
            reference.path,
            pending.filter((existing) => existing !== canonical),
          )
        }
        const unmerged = sheetUnmerges.get(reference.path) ?? new Set<string>()
        unmerged.add(canonical)
        sheetUnmerges.set(reference.path, unmerged)
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
      get tables(): readonly {
        readonly name: string
        readonly range: string
        readonly columns: readonly string[]
      }[] {
        const fromFile =
          sheetBytes === undefined ? [] : readTables(sheetBytes, reference.path, container)
        const pending = (sheetTables.get(reference.path) ?? []).map((spec) => ({
          name: spec.name,
          range: spec.ref,
          columns: spec.columns,
        }))
        return [...fromFile, ...pending]
      },
      addTable(
        range: string,
        options?: {
          readonly name?: string
          readonly columns?: readonly string[]
          readonly style?: string
        },
      ): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${range} cannot be made a table`,
            { ...at, reference: range },
          )
        }
        const planned = planTable(range, options, at, {
          existingRanges: [
            ...readTables(sheetBytes, reference.path, container).map((table) => table.range),
            ...(sheetTables.get(reference.path) ?? []).map((spec) => spec.ref),
          ],
          takenNames: takenTableNames(),
          headerTextAt: (headerRef) => headerTextOf(findCell(headerRef)),
        })
        for (const header of planned.headerCells) writeHeaderCell(header.ref, header.name)
        const specs = sheetTables.get(reference.path) ?? []
        specs.push(planned.spec)
        sheetTables.set(reference.path, specs)
      },
      addImage(image: Uint8Array, anchor: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so an image cannot be added`,
            { ...at, reference: anchor },
          )
        }
        const type = imageType(image)
        if (type === undefined) {
          throw new XlsxError('unwritable-value', 'Image is not a PNG, JPEG or GIF', {
            ...at,
            reference: anchor,
          })
        }
        const colon = anchor.indexOf(':')
        const first = parseWritableReference(colon === -1 ? anchor : anchor.slice(0, colon))
        const second = colon === -1 ? first : parseWritableReference(anchor.slice(colon + 1))
        const from = {
          column: Math.min(first.column, second.column) - 1,
          row: Math.min(first.row, second.row) - 1,
        }
        const to = {
          column: Math.max(first.column, second.column),
          row: Math.max(first.row, second.row),
        }
        const list = sheetImages.get(reference.path) ?? []
        list.push({ bytes: image, type, from, to })
        sheetImages.set(reference.path, list)
      },
      validate(range: string, rule: DataValidation): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${range} cannot be validated`,
            { ...at, reference: range },
          )
        }
        validations.add(reference.path, sqrefOf(range, at), at, rule)
      },
      clearValidations(): void {
        validations.clear(reference.path)
      },
      conditionalFormat(range: string, rule: ConditionalFormat): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${range} cannot be formatted`,
            { ...at, reference: range },
          )
        }
        const sqref = sqrefOf(range, at)
        const planned = planConditionalFormat(rule, sqref, at, dxfStyles)
        dxfStyles = planned.dxfStyles
        dxfColors.push(...planned.dxfColors)
        for (const spec of planned.specs) conditionalFormats.add(reference.path, spec)
      },
      clearConditionalFormats(): void {
        conditionalFormats.clear(reference.path)
      },
      comment(cellReference: string, text: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${cellReference} cannot be commented`,
            { ...at, reference: cellReference },
          )
        }
        const canonical = formatReference(parseWritableReference(cellReference))
        if (commentsRead.has(canonical)) {
          throw new XlsxError(
            'unsupported-edit',
            `Cell ${canonical} on sheet ${reference.name} already carries a comment, which this does not replace`,
            { ...at, reference: canonical },
          )
        }
        sheetRemovedComments.get(reference.path)?.delete(canonical)
        const notes = sheetComments.get(reference.path) ?? new Map<string, string>()
        notes.set(canonical, text)
        sheetComments.set(reference.path, notes)
      },
      removeComment(cellReference: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${cellReference} cannot be uncommented`,
            { ...at, reference: cellReference },
          )
        }
        const canonical = formatReference(parseWritableReference(cellReference))
        sheetComments.get(reference.path)?.delete(canonical)
        const removals = sheetRemovedComments.get(reference.path) ?? new Set<string>()
        removals.add(canonical)
        sheetRemovedComments.set(reference.path, removals)
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
      get pageSetup(): PageSetup {
        return page.mergedSetup(reference.path, sheetBytes)
      },
      setPageSetup(setup: PageSetup): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so its page setup cannot be set`,
            { ...at },
          )
        }
        page.setPageSetup(reference.path, setup, at)
      },
      get pageMargins(): PageMargins {
        return page.mergedMargins(reference.path, sheetBytes)
      },
      setPageMargins(margins: PageMargins): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so its margins cannot be set`,
            { ...at },
          )
        }
        page.setPageMargins(reference.path, margins, at)
      },
      get headerFooter(): HeaderFooter {
        return page.mergedHeaderFooter(reference.path, sheetBytes)
      },
      setHeaderFooter(headerFooter: HeaderFooter): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so its header and footer cannot be set`,
            { ...at },
          )
        }
        page.setHeaderFooter(reference.path, headerFooter)
      },
      get pageBreaks(): PageBreaks {
        return page.mergedBreaks(reference.path, sheetBytes)
      },
      addRowPageBreak(row: number): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so a page break cannot be added`,
            { ...at, reference: String(row) },
          )
        }
        page.addRowBreak(reference.path, row, at)
      },
      addColumnPageBreak(column: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so a page break cannot be added`,
            { ...at, reference: column },
          )
        }
        page.addColumnBreak(reference.path, column, at)
      },
      tabColor(color: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so its tab cannot be coloured`,
            { ...at, reference: color },
          )
        }
        sheetTabColors.set(reference.path, normalizeColor(color, at))
      },
      showGridlines(visible: boolean): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so its gridlines cannot be set`,
            { ...at },
          )
        }
        sheetGridlines.set(reference.path, visible)
      },
      showHeadings(visible: boolean): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so its headings cannot be set`,
            { ...at },
          )
        }
        sheetHeadings.set(reference.path, visible)
      },
      zoom(percent: number): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so it cannot be zoomed`,
            { ...at },
          )
        }
        if (!Number.isInteger(percent) || percent < 10 || percent > 400) {
          throw new XlsxError(
            'unwritable-value',
            `Zoom ${percent} is not a whole percentage between 10 and 400`,
            { ...at },
          )
        }
        sheetZoom.set(reference.path, percent)
      },
      groupRows(from: number, to: number, level = 1): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so rows cannot be grouped`,
            { ...at },
          )
        }
        if (!Number.isInteger(from) || from < 1 || !Number.isInteger(to) || to < from) {
          throw new XlsxError(
            'unwritable-value',
            `Rows ${from} to ${to} are not a forward range of row numbers`,
            { ...at },
          )
        }
        checkOutlineLevel(level, at)
        const levels = sheetRowGroups.get(reference.path) ?? new Map<number, number>()
        for (let row = from; row <= to; row++)
          levels.set(row, Math.max(levels.get(row) ?? 0, level))
        sheetRowGroups.set(reference.path, levels)
      },
      groupColumns(from: string, to: string, level = 1): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so columns cannot be grouped`,
            { ...at },
          )
        }
        const start = columnToIndex(from)
        const end = columnToIndex(to)
        if (end < start) {
          throw new XlsxError(
            'unwritable-value',
            `Columns ${from} to ${to} are not a forward range of columns`,
            { ...at },
          )
        }
        checkOutlineLevel(level, at)
        const levels = sheetColGroups.get(reference.path) ?? new Map<number, number>()
        for (let col = start; col <= end; col++)
          levels.set(col, Math.max(levels.get(col) ?? 0, level))
        sheetColGroups.set(reference.path, levels)
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
        refuseUnshiftable('its rows cannot be inserted into yet', undefined, SHIFTABLE_PARTS)
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
        refuseUnshiftable('its columns cannot be inserted into yet', before, SHIFTABLE_PARTS)
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
        refuseUnshiftable('its rows cannot be deleted yet', undefined, SHIFTABLE_PARTS)
        const spec = lineSpec('row', from, -count)
        const damage =
          deletionDamage(sheetBytes, spec) ??
          tableRowDamage(sheetBytes, reference.path, container, spec)
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
        refuseUnshiftable('its columns cannot be deleted yet', from, SHIFTABLE_PARTS)
        const spec = lineSpec('column', atColumn, -count)
        const resized = tableColumnDamage(sheetBytes, reference.path, container, spec)
        if (resized !== undefined) {
          throw new XlsxError(
            'unsupported-edit',
            `Sheet ${reference.name} carries ${resized}, whose columns cannot be deleted from yet`,
            { ...at, reference: from },
          )
        }
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
        sheetUnlinks.get(reference.path)?.delete(canonical)
        const links = sheetHyperlinks.get(reference.path) ?? new Map<string, Hyperlink>()
        links.set(canonical, target)
        sheetHyperlinks.set(reference.path, links)
      },
      unlink(cell: string): void {
        if (sheetBytes === undefined) {
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${cell} cannot be unlinked`,
            { ...at, reference: cell },
          )
        }
        const canonical = formatReference(parseWritableReference(cell))
        sheetHyperlinks.get(reference.path)?.delete(canonical)
        const unlinks = sheetUnlinks.get(reference.path) ?? new Set<string>()
        unlinks.add(canonical)
        sheetUnlinks.set(reference.path, unlinks)
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

  // A part path with a fresh trailing number, keeping the folder and extension, so
  // a duplicated `comments1.xml` becomes `comments7.xml` and collides with nothing.
  const freshPartPath = (original: string): string => {
    const taken = (candidate: string): boolean =>
      container.parts.has(candidate) || addedParts.has(candidate) || usedSheetPaths.has(candidate)
    const match = original.match(/^(.*?)(\d+)?(\.[^./]+)$/)
    const stem = match?.[1] ?? original
    const extension = match?.[3] ?? ''
    let n = 1
    while (taken(`${stem}${n}${extension}`)) n++
    return `${stem}${n}${extension}`
  }

  // The content type a part is declared with by an Override, or undefined when it
  // is covered by a Default (its extension), which the copy is covered by too.
  const overrideContentType = (path: string): string | undefined => {
    const types = partText(container, CONTENT_TYPES)
    const match = types?.match(
      new RegExp(
        `<Override PartName="/${path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*?ContentType="([^"]*)"`,
      ),
    )
    return match?.[1]
  }

  const copyWorksheet = (sourceName: string, newName: string): Worksheet => {
    const currentName = (ref: SheetRef): string => renames.get(ref.path) ?? ref.name
    const source = [...part.sheets, ...addedRefs.map((added) => added.reference)].find(
      (ref) => !removed.has(ref.path) && currentName(ref) === sourceName,
    )
    if (source === undefined) {
      throw new XlsxError('bad-reference', `No sheet named ${sourceName} to copy`, {})
    }
    const sourceBytes = container.parts.get(source.path) ?? addedSheets.get(source.path)
    if (sourceBytes === undefined) {
      throw new XlsxError('unsupported-edit', `Sheet ${sourceName} has no worksheet part to copy`, {
        part: source.path,
      })
    }
    checkSheetName(
      newName,
      sheets.map((sheet) => sheet.name),
    )

    let n = 1
    while (usedSheetPaths.has(`${workbookDir}worksheets/sheet${n}.xml`)) n++
    const newPath = `${workbookDir}worksheets/sheet${n}.xml`
    usedSheetPaths.add(newPath)

    // A sheet's own dependent parts are duplicated so the copy owns its own; parts
    // that need a unique name, id or shared media reworked are refused for now.
    const sourceRelsXml = partText(container, relationshipsPathFor(source.path))
    if (sourceRelsXml !== undefined) {
      let relsXml = sourceRelsXml
      for (const relationship of readRelationships(sourceRelsXml, source.path).values()) {
        if (relationship.external) continue
        const noun = relationship.type.endsWith('/table')
          ? 'a table'
          : relationship.type.endsWith('/pivotTable')
            ? 'a pivot table'
            : relationship.type.endsWith('/drawing')
              ? 'a drawing'
              : undefined
        if (noun !== undefined) {
          throw new XlsxError('unsupported-edit', `Cannot copy ${sourceName}: it carries ${noun}`, {
            part: source.path,
          })
        }
        const targetPath = resolveTarget(source.path, relationship.target)
        const bytes = container.parts.get(targetPath)
        if (bytes === undefined) continue
        const copyPath = freshPartPath(targetPath)
        addedParts.set(copyPath, bytes)
        const contentType = overrideContentType(targetPath)
        if (contentType !== undefined) addedOverrides.push({ path: copyPath, contentType })
        const oldFile = targetPath.slice(targetPath.lastIndexOf('/') + 1)
        const newFile = copyPath.slice(copyPath.lastIndexOf('/') + 1)
        relsXml = relsXml.replace(
          `Target="${relationship.target}"`,
          `Target="${relationship.target.replace(oldFile, newFile)}"`,
        )
      }
      addedParts.set(relationshipsPathFor(newPath), new TextEncoder().encode(relsXml))
    }

    const reference: SheetRef = {
      name: newName,
      path: newPath,
      sheetId: String(++maxSheetId),
      state: 'visible',
    }
    addedSheets.set(newPath, sourceBytes)
    addedRefs.push({
      reference,
      relationshipId: `rId${++maxRelationshipId}`,
      target: newPath.slice(workbookDir.length),
    })
    const sheet = makeWorksheet(reference)
    sheets.push(sheet)
    return sheet
  }

  // The part-level features that each write their own parts over the draft. One
  // entry per feature is the whole registration: the write pass reads `pending`
  // for its "anything to do?" gate and calls `contribute` to do the work, so a new
  // contributor is a single list entry rather than a probe and a call kept in step.
  // They run before the line-ops shift so a reference they write moves with it, and
  // none touches another's parts, so their order among themselves is free.
  const contributors: ReadonlyArray<{
    pending: () => boolean
    contribute: (draft: ContainerDraft) => void
  }> = [
    {
      pending: () => sheetHyperlinks.size > 0 || sheetUnlinks.size > 0,
      contribute: (draft) => contributeHyperlinks(draft, sheetHyperlinks, sheetUnlinks, removed),
    },
    {
      pending: () => sheetComments.size > 0 || sheetRemovedComments.size > 0,
      contribute: (draft) =>
        contributeComments(draft, sheetComments, sheetRemovedComments, removed),
    },
    {
      pending: () => sheetImages.size > 0,
      contribute: (draft) => contributeImages(draft, sheetImages, removed),
    },
    {
      pending: () => sheetTables.size > 0,
      contribute: (draft) => contributeTables(draft, sheetTables, removed),
    },
  ]

  const toBytes = (): Uint8Array => {
    // A change carries new bytes for a part; null deletes it. Every part not
    // named here is passed through still compressed, never inflated or rebuilt.
    const changes = new Map<string, Uint8Array | null>()
    // Reads and stages parts over the container: the plumbing the part-level
    // features (comments here, images and tables below) share.
    const draft = createContainerDraft(container, changes)
    // Parts a copySheet duplicated (dependent parts and the copy's own rels) are
    // written verbatim; the copy's worksheet part flows through the patch below.
    for (const [path, bytes] of addedParts) changes.set(path, bytes)
    // One source for "is anything pending?". The sheet-path maps come from
    // patchInputs, so a new kind of sheet edit registers only there; a part- or
    // workbook-level container patchSheet does not take adds one line here. A
    // format() with no set() records a style override and no value edit; a
    // protect(), merge() or setRowHeight() records neither and still rewrites.
    const pendingEdits: ReadonlyArray<() => boolean> = [
      ...patchInputs.map((map) => () => map.size > 0),
      ...contributors.map((contributor) => contributor.pending),
      () => validations.cleared.size > 0,
      () => conditionalFormats.cleared.size > 0,
      () => page.hasPending(),
      () => addedRefs.length > 0,
      () => renames.size > 0,
      () => removed.size > 0,
      () => pendingNames.size > 0,
      () => removedNames.size > 0,
      () => sheetPendingNames.size > 0,
      () => sheetRemovedNames.size > 0,
      () => lineOps.length > 0,
      () => Object.keys(pendingProperties).length > 0,
      () => sheetStates.size > 0,
    ]
    if (!pendingEdits.some((has) => has())) {
      return container.write(changes)
    }

    const encoder = new TextEncoder()

    // Excel rebuilds the calculation chain, but a stale one makes it offer to
    // repair the file. Dropping it also touches the workbook rels and the content
    // types, which the added-sheet wiring below writes in the same pass.
    const hadCalcChain = container.parts.has(CALCULATION_CHAIN)
    if (hadCalcChain) changes.set(CALCULATION_CHAIN, null)

    // set() already assigned every style index in memory; here the tables are
    // spliced into the original once, then the low-volume dxf highlights are
    // folded on — dxfs sit in their own table, so replaying them onto the
    // serialized styles reproduces the same bytes the threaded string held.
    if (session !== undefined && (session.changed() || dxfColors.length > 0)) {
      let styles = session.serialize()
      for (const color of dxfColors) styles = ensureDxf(styles, color).xml
      if (styles !== stylesXml) changes.set('xl/styles.xml', encoder.encode(styles))
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
      ...validations.cleared,
      ...conditionalFormats.cleared,
      ...addedSheets.keys(),
    ])) {
      if (removed.has(path)) continue
      let bytes = container.parts.get(path) ?? addedSheets.get(path)
      if (bytes === undefined) continue
      // Clears drop the file's elements before the patch adds this session's, so a
      // clear-then-add ends with only the added rules.
      if (validations.cleared.has(path)) {
        bytes = encoder.encode(withoutDataValidations(decodeXmlPart(bytes, path)))
      }
      if (conditionalFormats.cleared.has(path)) {
        bytes = encoder.encode(withoutConditionalFormatting(decodeXmlPart(bytes, path)))
      }
      const pending = edits.get(path) ?? EMPTY_EDITS
      const at: SheetLocation = {
        sheet:
          part.sheets.find((s) => s.path === path)?.name ??
          addedRefs.find((added) => added.reference.path === path)?.reference.name,
        part: path,
      }
      const patched = patchSheet(bytes, pending, date1904, indexes, styleOverrides.get(path), at, {
        protection: sheetProtections.get(path),
        merges: sheetMerges.get(path),
        rowHeights: sheetRowHeights.get(path),
        columnWidths: sheetColumnWidths.get(path),
        autoFilter: sheetAutoFilters.get(path),
        freeze: sheetFreezes.get(path),
        hiddenRows: sheetHiddenRows.get(path),
        hiddenColumns: sheetHiddenColumns.get(path),
        tabColor: sheetTabColors.get(path),
        showGridLines: sheetGridlines.get(path),
        showRowColHeaders: sheetHeadings.get(path),
        zoomScale: sheetZoom.get(path),
        rowOutlineLevels: sheetRowGroups.get(path),
        colOutlineLevels: sheetColGroups.get(path),
        dataValidations: validations.pending.get(path),
        conditionalFormats: conditionalFormats.pending.get(path),
      })
      // Unmerge runs after the merge-add above, on the written sheet, so a range
      // both merged and unmerged this session ends up gone.
      const unmerges = sheetUnmerges.get(path)
      changes.set(
        path,
        unmerges === undefined || unmerges.size === 0
          ? patched
          : encoder.encode(withoutMergeCells(decodeXmlPart(patched, path), unmerges)),
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

    // Composes a top-level part written once this pass — the workbook, its rels,
    // the content types — from its current text, and writes it back only when the
    // transform changed something, so an untouched part is never rewritten.
    const rewritePart = (path: string, transform: (xml: string) => string): void => {
      const xml = partText(container, path)
      if (xml === undefined) return
      const updated = transform(xml)
      if (updated !== xml) changes.set(path, encoder.encode(updated))
    }

    // Each part-level feature writes its parts and declares their content types on
    // the draft. They run before the line-ops shift so a reference one writes (a
    // hyperlink's cell) moves with it.
    for (const contributor of contributors) contributor.contribute(draft)

    // A copySheet duplicated dependent parts whose bytes are already staged; their
    // content types declare through the same draft as the contributors.
    for (const override of addedOverrides) {
      draft.declareOverride(override.path, override.contentType)
    }

    // Page setup writes elements the schema orders margins, setup, headerFooter,
    // rowBreaks, colBreaks. Applying them in that order lands each in the right
    // place whether the sheet already had it or it is inserted fresh here.
    for (const path of page.paths()) {
      if (removed.has(path)) continue
      const sheetXml = draft.text(path)
      if (sheetXml === undefined) continue
      changes.set(path, encoder.encode(page.apply(sheetXml, path)))
    }

    // applyLineShifts moves references across the workbook for this session row and
    // column inserts and deletes, and returns the defined names shifted the same way.
    const namesToWrite = applyLineShifts(draft, container, changes, {
      sheets: part.sheets,
      addedSheetPaths: addedSheets.keys(),
      removed,
      lineOps,
      fileNames,
      pendingNames,
      date1904,
    })

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

    // A scoped name carries the sheet's index in the final <sheets> order, which
    // is the file's own minus removals then the added ones appended — how
    // withSheetRemoved and withSheetsAdded reorder the element.
    const finalSheetPaths = [
      ...part.sheets.filter((sheet) => !removed.has(sheet.path)).map((sheet) => sheet.path),
      ...renamedAdded.map((added) => added.reference.path),
    ]
    const scopedNameWrites: SheetScopedName[] = []
    for (const [path, pending] of sheetPendingNames) {
      const localSheetId = finalSheetPaths.indexOf(path)
      if (localSheetId < 0) continue
      for (const [name, refersTo] of pending) {
        scopedNameWrites.push({ name, localSheetId, refersTo })
      }
    }
    const scopedNameRemovals: { name: string; localSheetId: number }[] = []
    for (const [path, removals] of sheetRemovedNames) {
      const localSheetId = finalSheetPaths.indexOf(path)
      if (localSheetId < 0) continue
      for (const name of removals) scopedNameRemovals.push({ name, localSheetId })
    }
    rewritePart(part.path, (workbookXml) => {
      let updated = wroteFormula ? withRecalculation(workbookXml) : workbookXml
      for (const [path, name] of renames) {
        const original = originalName(path)
        if (original !== undefined && !removed.has(path)) {
          updated = withSheetRenamed(updated, original, name)
        }
      }
      updated = withSheetsAdded(updated, renamedAdded)
      // Visibility lands by the sheet's current name — after a rename, and once an
      // added sheet's element exists to carry it.
      for (const [path, state] of sheetStates) {
        if (removed.has(path)) continue
        const name =
          renames.get(path) ??
          originalName(path) ??
          renamedAdded.find((added) => added.reference.path === path)?.reference.name
        if (name !== undefined) updated = withSheetState(updated, name, state)
      }
      for (const path of removedExisting) {
        const original = originalName(path)
        if (original !== undefined) updated = withSheetRemoved(updated, original)
      }
      return withDefinedNames(updated, namesToWrite, removedNames, {
        write: scopedNameWrites,
        remove: scopedNameRemovals,
      })
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

    // Core properties splice into their part, or open one wired to the package
    // root and declared in the content types when the file carries none.
    let createdCoreProperties = false
    if (Object.keys(pendingProperties).length > 0) {
      const existing = partText(container, CORE_PROPERTIES_PART)
      changes.set(
        CORE_PROPERTIES_PART,
        encoder.encode(writeCoreProperties(existing, pendingProperties)),
      )
      if (existing === undefined) {
        createdCoreProperties = true
        rewritePart(
          ROOT_RELATIONSHIPS,
          (relsXml) =>
            withRelationship(relsXml, CORE_PROPERTIES_RELATIONSHIP, CORE_PROPERTIES_PART).xml,
        )
      }
    }

    rewritePart(CONTENT_TYPES, (contentTypesXml) => {
      let updated = withSheetContentTypes(
        hadCalcChain ? withoutOverride(contentTypesXml, CALCULATION_CHAIN) : contentTypesXml,
        addedRefs,
      )
      for (const path of removedExisting) updated = withoutOverride(updated, path)
      if (createdCoreProperties) {
        updated = withContentTypeOverride(
          updated,
          CORE_PROPERTIES_PART,
          CORE_PROPERTIES_CONTENT_TYPE,
        )
      }
      return draft.applyContentTypes(updated)
    })

    return container.write(changes)
  }

  return {
    sheets,
    sheet: (name: string) => sheets.find((candidate) => candidate.name === name),
    addSheet: addWorksheet,
    copySheet: copyWorksheet,
    get definedNames(): ReadonlyMap<string, string> {
      const all = new Map([...fileNames, ...pendingNames])
      for (const name of removedNames) all.delete(name)
      return all
    },
    defineName(name: string, refersTo: string): void {
      checkDefinedName(name, refersTo)
      removedNames.delete(name)
      pendingNames.set(name, refersTo)
    },
    removeDefinedName(name: string): void {
      pendingNames.delete(name)
      removedNames.add(name)
    },
    get properties(): DocumentProperties {
      const fileXml = partText(container, CORE_PROPERTIES_PART)
      const file = fileXml === undefined ? {} : readCoreProperties(fileXml)
      return { ...file, ...pendingProperties }
    },
    setProperties(properties: DocumentProperties): void {
      pendingProperties = { ...pendingProperties, ...properties }
    },
    epoch: date1904 ? 1904 : 1900,
    resolveColor: (color: Color) => resolveColor(color, themePalette),
    toBytes,
  }
}
