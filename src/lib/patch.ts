import { dateToSerial } from './date.js'
import { XlsxError } from './errors.js'
import {
  LAST_COLUMN,
  LAST_ROW,
  canonicalReference,
  formatReference,
  parseFileReference,
  parseReference,
} from './reference.js'
import { findUnwritableCharacter, readXmlBytes, withAttribute } from './xml.js'

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

/**
 * What a caller may still do once the worksheet is protected. Each is `true` when
 * the action stays permitted; omitted falls to Excel's default — the format,
 * insert, delete, sort and filter family locked, and selecting cells allowed.
 * Cells resist editing per their own `locked` flag, which only bites here.
 */
export interface SheetProtection {
  readonly selectLockedCells?: boolean
  readonly selectUnlockedCells?: boolean
  readonly formatCells?: boolean
  readonly formatColumns?: boolean
  readonly formatRows?: boolean
  readonly insertColumns?: boolean
  readonly insertRows?: boolean
  readonly insertHyperlinks?: boolean
  readonly deleteColumns?: boolean
  readonly deleteRows?: boolean
  readonly sort?: boolean
  readonly autoFilter?: boolean
  readonly pivotTables?: boolean
  /** Editing drawing objects; Excel locks this by default when protecting. */
  readonly editObjects?: boolean
  /** Editing scenarios; Excel locks this by default when protecting. */
  readonly editScenarios?: boolean
}

// Each permission maps to a CT_SheetProtection attribute where 1 means the action
// is locked and 0 means it is allowed. Only the ones the caller names are written;
// the rest keep the schema defaults.
const PROTECTION_PERMISSIONS: ReadonlyArray<readonly [keyof SheetProtection, string]> = [
  ['selectLockedCells', 'selectLockedCells'],
  ['selectUnlockedCells', 'selectUnlockedCells'],
  ['formatCells', 'formatCells'],
  ['formatColumns', 'formatColumns'],
  ['formatRows', 'formatRows'],
  ['insertColumns', 'insertColumns'],
  ['insertRows', 'insertRows'],
  ['insertHyperlinks', 'insertHyperlinks'],
  ['deleteColumns', 'deleteColumns'],
  ['deleteRows', 'deleteRows'],
  ['sort', 'sort'],
  ['autoFilter', 'autoFilter'],
  ['pivotTables', 'pivotTables'],
]

function sheetProtectionElement(protection: SheetProtection, prefix: string): string {
  // objects and scenarios default to unlocked in the schema, but Excel's Protect
  // Sheet locks them, so they are always written, defaulting to locked.
  let attributes =
    ` sheet="1" objects="${protection.editObjects === true ? '0' : '1'}"` +
    ` scenarios="${protection.editScenarios === true ? '0' : '1'}"`
  for (const [key, attribute] of PROTECTION_PERMISSIONS) {
    const permitted = protection[key]
    if (permitted !== undefined) attributes += ` ${attribute}="${permitted ? '0' : '1'}"`
  }
  return `<${prefix}sheetProtection${attributes}/>`
}

/** 0 means the action stays permitted, 1 that it is locked; anything else off. */
const isPermitted = (value: string): boolean => value === '0' || value === 'false'

/**
 * The protection a sheet declares, in the same shape `protect()` takes, or
 * undefined when it is not protected. Only the attributes the element carries
 * are reported; the rest keep their defaults.
 */
export function readSheetProtection(bytes: Uint8Array): SheetProtection | undefined {
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'sheetProtection') continue
    if (!isProtected(event.attributes.get('sheet'))) return undefined

    const result: { -readonly [K in keyof SheetProtection]?: boolean } = {}
    for (const [key, attribute] of PROTECTION_PERMISSIONS) {
      const value = event.attributes.get(attribute)
      if (value !== undefined) result[key] = isPermitted(value)
    }
    const objects = event.attributes.get('objects')
    if (objects !== undefined) result.editObjects = isPermitted(objects)
    const scenarios = event.attributes.get('scenarios')
    if (scenarios !== undefined) result.editScenarios = isPermitted(scenarios)
    return result
  }
  return undefined
}

const isProtected = (value: string | undefined): boolean => value === '1' || value === 'true'

/** A row's stored height in points, keyed by its one-based number. */
export function readRowHeights(bytes: Uint8Array): ReadonlyMap<number, number> {
  const heights = new Map<number, number>()
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'row') continue
    const stored = event.attributes.get('ht')
    if (stored === undefined) continue
    const number = Number(event.attributes.get('r'))
    const height = Number(stored)
    if (Number.isInteger(number) && number >= 1 && Number.isFinite(height)) {
      heights.set(number, height)
    }
  }
  return heights
}

/**
 * The column-width ranges a sheet stores, each width covering columns `min` to
 * `max`. Kept as ranges rather than expanded per column, since one `<col>` may
 * span the whole sheet.
 */
export function readColumnWidths(
  bytes: Uint8Array,
): readonly { min: number; max: number; width: number }[] {
  const ranges: { min: number; max: number; width: number }[] = []
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'col') continue
    const stored = event.attributes.get('width')
    if (stored === undefined) continue
    const min = Number(event.attributes.get('min'))
    const max = Number(event.attributes.get('max'))
    const width = Number(stored)
    if (Number.isInteger(min) && Number.isInteger(max) && Number.isFinite(width)) {
      ranges.push({ min, max, width })
    }
  }
  return ranges
}

/** Element content only. Quotes need no escaping there, and Excel leaves them. */
const escapeXml = (text: string) =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // XML 1.0 has a parser fold a literal CR, and the CR of a CRLF, into a bare
    // LF before the document reaches the application. Written raw it is lost.
    .replace(/\r/g, '&#13;')

/**
 * A formula, or a refusal. Types stop a TypeScript caller passing anything
 * else; a JS caller, a JSON payload or an `any` at a boundary reach here, and
 * reading `.formula` off them threw outside the error contract.
 */
function formulaOf(value: object, reference: string, at: SheetLocation): string {
  if ('formula' in value && typeof value.formula === 'string') return value.formula
  throw new XlsxError(
    'unwritable-value',
    `Cell ${reference} was given an object that names no string formula`,
    { ...at, reference },
  )
}

/**
 * Every refusal writing a cell can make, so `set()` can make it at the call
 * instead of leaving a workbook that only fails once it is saved. Takes
 * `unknown` because it is the boundary the value has to get past.
 */
export function checkWritable(
  reference: string,
  value: unknown,
  date1904: boolean,
  at: SheetLocation = {},
): void {
  if (value === null || typeof value === 'boolean') return

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new XlsxError('unwritable-value', `Cell ${reference} cannot hold ${value}`, {
        ...at,
        reference,
      })
    }
    return
  }

  if (value instanceof Date) {
    dateToSerial(value, date1904, { ...at, reference })
    return
  }

  if (typeof value !== 'string' && (typeof value !== 'object' || value === null)) {
    throw new XlsxError(
      'unwritable-value',
      `Cell ${reference} cannot hold a value of type ${typeof value}`,
      { ...at, reference },
    )
  }

  const unwritable = findUnwritableCharacter(
    typeof value === 'string' ? value : formulaOf(value, reference, at),
  )
  if (unwritable !== undefined) {
    throw new XlsxError(
      'unwritable-value',
      `Cell ${reference} holds ${unwritable}, which cannot be written to xml`,
      { ...at, reference },
    )
  }
}

/**
 * Refuses a `protect()` argument that is not an options object, at the call.
 * Types stop a TypeScript caller; a JS caller or an `any` reaches here, where
 * reading a flag off null crashed at save time and took the whole batch down.
 */
export function checkProtection(options: unknown, at: SheetLocation = {}): void {
  if (typeof options !== 'object' || options === null) {
    throw new XlsxError('unwritable-value', 'Sheet protection options must be an object', { ...at })
  }
}

/**
 * A formula whose result covers more cells than the one holding it. The others
 * carry a cached value and no formula of their own, so replacing this cell
 * leaves them owned by nothing.
 */
export interface SpillingFormula {
  readonly kind: 'shared' | 'array' | 'dataTable'
  /** The si for a shared formula, the covered range for the other two. */
  readonly name: string
}

/** A merged region, as a rectangle so a whole-column merge costs no per-cell store. */
interface MergeRange {
  /** Canonical top-left, the one member a written value would actually show in. */
  readonly anchor: string
  readonly minRow: number
  readonly maxRow: number
  readonly minColumn: number
  readonly maxColumn: number
}

function parseMergeRange(ref: string): MergeRange | undefined {
  const colon = ref.indexOf(':')
  if (colon === -1) return undefined
  const start = parseReference(ref.slice(0, colon))
  const end = parseReference(ref.slice(colon + 1))
  const minRow = Math.min(start.row, end.row)
  const minColumn = Math.min(start.column, end.column)
  return {
    anchor: formatReference({ row: minRow, column: minColumn }),
    minRow,
    maxRow: Math.max(start.row, end.row),
    minColumn,
    maxColumn: Math.max(start.column, end.column),
  }
}

const canonicalMerge = (merge: MergeRange): string =>
  `${merge.anchor}:${formatReference({ row: merge.maxRow, column: merge.maxColumn })}`

/**
 * The canonical `A1:B2` form of a range to merge, so `b2:a1` and `A1:B2` are the
 * same merge. Refuses anything that is not two references either side of a colon.
 */
export function mergeRangeReference(range: unknown, at: SheetLocation = {}): string {
  if (typeof range !== 'string') {
    throw new XlsxError('bad-reference', `A merge range must be a string, not ${typeof range}`, {
      ...at,
      reference: String(range),
    })
  }
  const parsed = parseMergeRange(range)
  if (parsed === undefined) {
    throw new XlsxError('bad-reference', `"${range}" is not an A1:B2 style range to merge`, {
      ...at,
      reference: range,
    })
  }
  // A caller's merge, like a caller's cell, must name cells the sheet can hold;
  // parseMergeRange itself stays lenient because it also parses merges read from
  // a file, where an odd range is left alone rather than refused.
  if (
    parsed.minRow < 1 ||
    parsed.maxRow > LAST_ROW ||
    parsed.minColumn < 1 ||
    parsed.maxColumn > LAST_COLUMN
  ) {
    throw new XlsxError('bad-reference', `"${range}" is outside the sheet`, {
      ...at,
      reference: range,
    })
  }
  return canonicalMerge(parsed)
}

/** All keyed by canonical reference, so `a1` and `$A$1` find the same cell. */
export interface SheetIndex {
  /** The style index a cell carried when the file was read. */
  readonly styles: ReadonlyMap<string, number>
  /** Cells that other cells depend on for their formula. */
  readonly sharedFormulas: ReadonlyMap<string, SpillingFormula>
  readonly merges: readonly MergeRange[]
}

/** One pass, because set() needs all of these before it accepts an edit. */
export function indexSheet(bytes: Uint8Array): SheetIndex {
  const styles = new Map<string, number>()
  const sharedFormulas = new Map<string, SpillingFormula>()
  const shape = readShape(bytes)

  for (const row of shape.rows) {
    for (const cell of row.cells) {
      if (cell.style === undefined && cell.spillingFormula === undefined) continue
      // A cell no column letter can name is one set() could never target either,
      // so it is skipped rather than allowed to throw the whole index away.
      const reference = canonicalReference({ row: row.row, column: cell.column })
      if (reference === undefined) continue
      if (cell.style !== undefined) styles.set(reference, Number(cell.style))
      if (cell.spillingFormula !== undefined) {
        sharedFormulas.set(reference, cell.spillingFormula)
      }
    }
  }

  return { styles, sharedFormulas, merges: shape.merges }
}

/** The anchor of a merge `reference` is buried inside, or undefined if it is free. */
export function mergeAnchorFor(index: SheetIndex, reference: string): string | undefined {
  const { row, column } = parseReference(reference)
  for (const merge of index.merges) {
    const inside =
      row >= merge.minRow &&
      row <= merge.maxRow &&
      column >= merge.minColumn &&
      column <= merge.maxColumn
    if (inside && merge.anchor !== reference) return merge.anchor
  }
  return undefined
}

export function sharedFormulaRefusal(
  reference: string,
  master: SpillingFormula,
  at: SheetLocation = {},
): XlsxError {
  const what =
    master.kind === 'shared'
      ? `defines shared formula ${master.name}`
      : `holds the ${master.kind === 'array' ? 'array formula' : 'data table'} covering ${master.name}`
  return new XlsxError(
    'unwritable-value',
    `Cell ${reference} ${what}; overwriting it would break the cells that depend on it`,
    { ...at, reference },
  )
}

export function mergeRefusal(reference: string, anchor: string, at: SheetLocation = {}): XlsxError {
  return new XlsxError(
    'unwritable-value',
    `Cell ${reference} is merged into ${anchor}; a value there would never show, so write ${anchor} instead`,
    { ...at, reference },
  )
}

/**
 * An existing style is carried over so formatting survives an edit, which means
 * a Date written into a cell with no date format will show as a number.
 */
function cellElement(
  reference: string,
  value: CellInput,
  style: string | undefined,
  date1904: boolean,
  sharedStrings: ReadonlyMap<string, number> | undefined,
  prefix: string,
  at: SheetLocation,
): string {
  checkWritable(reference, value, date1904, at)

  const attributes = style === undefined ? '' : ` s="${style}"`
  const c = `${prefix}c`
  const v = `${prefix}v`

  if (value === null) return `<${c} r="${reference}"${attributes}/>`

  if (typeof value === 'object' && !(value instanceof Date)) {
    // No cached result: nothing here computes one, and a stale one is worse.
    const f = `${prefix}f`
    return `<${c} r="${reference}"${attributes}><${f}>${escapeXml(value.formula)}</${f}></${c}>`
  }
  if (typeof value === 'string') {
    const shared = sharedStrings?.get(value)
    if (shared !== undefined) {
      return `<${c} r="${reference}"${attributes} t="s"><${v}>${shared}</${v}></${c}>`
    }
    const space = value === value.trim() ? '' : ' xml:space="preserve"'
    return (
      `<${c} r="${reference}"${attributes} t="inlineStr">` +
      `<${prefix}is><${prefix}t${space}>${escapeXml(value)}</${prefix}t></${prefix}is></${c}>`
    )
  }
  if (typeof value === 'boolean') {
    return `<${c} r="${reference}"${attributes} t="b"><${v}>${value ? 1 : 0}</${v}></${c}>`
  }
  if (value instanceof Date) {
    return `<${c} r="${reference}"${attributes}><${v}>${dateToSerial(value, date1904)}</${v}></${c}>`
  }
  return `<${c} r="${reference}"${attributes}><${v}>${value}</${v}></${c}>`
}

interface CellSpan {
  readonly column: number
  readonly start: number
  readonly end: number
  readonly style: string | undefined
  /** Set when other cells take their formula from this one. */
  readonly spillingFormula: SpillingFormula | undefined
}

interface RowSpan {
  readonly row: number
  readonly start: number
  /** Offset just after the row's open tag, where its attributes end. */
  readonly openEnd: number
  /** Offset just before `</row>`, where a new cell is appended. */
  readonly contentEnd: number
  /** A self closing row holds nothing, so it is rewritten rather than spliced into. */
  readonly selfClosing: boolean
  readonly end: number
  readonly cells: CellSpan[]
}

interface DimensionSpan {
  readonly start: number
  readonly end: number
  readonly ref: string
}

interface SheetShape {
  /** Namespace prefix the document writes its elements with, `x:` or empty. */
  readonly prefix: string
  readonly dimension: DimensionSpan | undefined
  readonly rows: RowSpan[]
  readonly merges: MergeRange[]
  /** Offset just before `</sheetData>`, where a new row is appended. */
  readonly contentEnd: number
  readonly selfClosing: boolean
  readonly dataStart: number
  readonly dataEnd: number
  /** An existing sheetProtection element, replaced in place rather than doubled. */
  readonly protection: { start: number; end: number } | undefined
  /** An existing autoFilter element, replaced in place rather than doubled. */
  readonly autoFilter: { start: number; end: number } | undefined
  /** The first sheetView, where a freeze pane is placed or replaces its own. */
  readonly sheetView: { start: number; end: number; selfClosing: boolean } | undefined
  /** An existing pane inside that sheetView, replaced rather than doubled. */
  readonly pane: { start: number; end: number } | undefined
  /** End of the worksheet root's open tag, where a fresh sheetPr is anchored. */
  readonly worksheetStart: number
  /** An existing sheetPr, so a tab colour joins it rather than opening a second. */
  readonly sheetPr: { start: number; end: number; selfClosing: boolean } | undefined
  /** An existing tabColor inside sheetPr, replaced in place. */
  readonly tabColor: { start: number; end: number } | undefined
  /** An existing sheetFormatPr, whose outline-level hints a grouping updates. */
  readonly sheetFormatPr: { start: number; end: number } | undefined
  /** An existing dataValidations, so a new rule joins it and bumps its count. */
  readonly dataValidations:
    | { openStart: number; openEnd: number; insertAt: number; selfClosing: boolean; count: number }
    | undefined
  /** Start of the first sibling that must follow dataValidations, or -1. */
  readonly laterSiblingStart: number
  /** Start of the worksheet's close tag, the fallback insertion point. */
  readonly worksheetEnd: number
  /** The highest cfRule priority the sheet uses, so a new rule can outrank it. */
  readonly maxPriority: number
  /** An existing mergeCells element, so a new merge joins it rather than a second. */
  readonly mergeContainer:
    | { openStart: number; openEnd: number; insertAt: number; selfClosing: boolean; count: number }
    | undefined
  /** An existing cols element, before sheetData, that a width joins. */
  readonly colContainer:
    | { openStart: number; openEnd: number; insertAt: number; selfClosing: boolean }
    | undefined
  /** The col entries, so a width lands in the range covering its column. */
  readonly cols: { start: number; end: number; min: number; max: number }[]
}

function readShape(bytes: Uint8Array, at: SheetLocation = {}): SheetShape {
  const rows: RowSpan[] = []
  const merges: MergeRange[] = []
  let dimension: DimensionSpan | undefined
  let contentEnd = -1
  let dataStart = -1
  let dataEnd = -1
  let selfClosing = false
  let protection: { start: number; end: number } | undefined
  let autoFilter: { start: number; end: number } | undefined
  let openAutoFilter = -1
  let sheetView: { start: number; end: number; selfClosing: boolean } | undefined
  let pane: { start: number; end: number } | undefined
  let worksheetStart = -1
  let sheetPr: { start: number; end: number; selfClosing: boolean } | undefined
  let tabColor: { start: number; end: number } | undefined
  let sheetFormatPr: { start: number; end: number } | undefined
  let dvOpenStart = -1
  let dvOpenEnd = -1
  let dvInsertAt = -1
  let dvSelfClosing = false
  let dvCount = 0
  let laterSiblingStart = -1
  let worksheetEnd = -1
  let maxPriority = 0
  let mergeOpenStart = -1
  let mergeOpenEnd = -1
  let mergeInsertAt = -1
  let mergeSelfClosing = false
  let mergeCount = 0
  let colOpenStart = -1
  let colOpenEnd = -1
  let colInsertAt = -1
  let colSelfClosing = false
  const cols: { start: number; end: number; min: number; max: number }[] = []

  let prefix = ''
  let lastRow = 0
  let openDimension = -1
  let currentRow: { row: number; start: number; openEnd: number } | undefined
  let currentCells: CellSpan[] = []
  let cellStart = -1
  let cellColumn = 0
  let cellStyle: string | undefined
  let master: SpillingFormula | undefined
  let openCell = false

  for (const event of readXmlBytes(bytes)) {
    if (event.kind === 'open') {
      if (event.localName === 'dimension') {
        const ref = event.attributes.get('ref')
        if (ref !== undefined) {
          dimension = { start: event.start, end: event.end, ref }
          openDimension = event.selfClosing ? -1 : event.start
        }
        continue
      }
      if (event.localName === 'sheetData') {
        const colon = event.name.indexOf(':')
        prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
        dataStart = event.start
        selfClosing = event.selfClosing
        if (selfClosing) {
          dataEnd = event.end
          contentEnd = event.end
        }
        continue
      }
      if (event.localName === 'row') {
        // A row without r is the one after the row before it, which is not the
        // same as the count of rows seen once any row declares its number.
        const declared = event.attributes.get('r')
        if (declared === undefined) lastRow += 1
        else {
          lastRow = Number(declared)
          if (!Number.isInteger(lastRow)) {
            throw new XlsxError(
              'invalid-content',
              `Row number "${declared}" in the file is not a valid row`,
              { ...at },
            )
          }
        }
        currentRow = { row: lastRow, start: event.start, openEnd: event.end }
        currentCells = []
        cellColumn = 0
        if (event.selfClosing) {
          rows.push({
            ...currentRow,
            contentEnd: event.end,
            end: event.end,
            selfClosing: true,
            cells: [],
          })
          currentRow = undefined
        }
        continue
      }
      if (event.localName === 'c') {
        const reference = event.attributes.get('r')
        cellColumn =
          reference === undefined ? cellColumn + 1 : parseFileReference(reference, at).column
        cellStyle = event.attributes.get('s')
        cellStart = event.start
        master = undefined
        openCell = !event.selfClosing
        if (event.selfClosing) {
          currentCells.push({
            column: cellColumn,
            start: event.start,
            end: event.end,
            style: cellStyle,
            spillingFormula: undefined,
          })
        }
      }
      if (event.localName === 'cols') {
        colOpenStart = event.start
        colOpenEnd = event.end
        colSelfClosing = event.selfClosing
        if (event.selfClosing) colInsertAt = event.end
        continue
      }
      if (event.localName === 'col') {
        const min = Number(event.attributes.get('min'))
        const max = Number(event.attributes.get('max'))
        if (Number.isInteger(min) && Number.isInteger(max)) {
          cols.push({ start: event.start, end: event.end, min, max })
        }
        continue
      }
      if (event.localName === 'mergeCells') {
        mergeOpenStart = event.start
        mergeOpenEnd = event.end
        mergeSelfClosing = event.selfClosing
        if (event.selfClosing) mergeInsertAt = event.end
        continue
      }
      if (event.localName === 'mergeCell') {
        mergeCount++
        const ref = event.attributes.get('ref')
        const merge = ref === undefined ? undefined : parseMergeRange(ref)
        if (merge !== undefined) merges.push(merge)
        continue
      }
      if (event.localName === 'sheetProtection') {
        protection = { start: event.start, end: event.end }
        continue
      }
      if (event.localName === 'autoFilter') {
        // A filtered autoFilter carries children, so its close tag ends the span.
        if (event.selfClosing) autoFilter = { start: event.start, end: event.end }
        else openAutoFilter = event.start
        continue
      }
      if (event.localName === 'sheetView' && sheetView === undefined) {
        sheetView = { start: event.start, end: event.end, selfClosing: event.selfClosing }
        continue
      }
      if (event.localName === 'pane' && pane === undefined) {
        pane = { start: event.start, end: event.end }
        continue
      }
      if (event.localName === 'worksheet' && worksheetStart === -1) {
        worksheetStart = event.end
        continue
      }
      if (event.localName === 'sheetPr' && sheetPr === undefined) {
        sheetPr = { start: event.start, end: event.end, selfClosing: event.selfClosing }
        continue
      }
      if (event.localName === 'tabColor' && tabColor === undefined) {
        tabColor = { start: event.start, end: event.end }
        continue
      }
      if (event.localName === 'sheetFormatPr' && sheetFormatPr === undefined) {
        sheetFormatPr = { start: event.start, end: event.end }
        continue
      }
      if (event.localName === 'dataValidations' && dvOpenStart === -1) {
        dvOpenStart = event.start
        dvOpenEnd = event.end
        dvSelfClosing = event.selfClosing
        if (event.selfClosing) dvInsertAt = event.end
        continue
      }
      if (event.localName === 'dataValidation') {
        dvCount += 1
        continue
      }
      if (AFTER_DATA_VALIDATIONS.has(event.localName) && laterSiblingStart === -1) {
        laterSiblingStart = event.start
        continue
      }
      if (event.localName === 'cfRule') {
        const priority = Number(event.attributes.get('priority'))
        if (Number.isFinite(priority)) maxPriority = Math.max(maxPriority, priority)
        continue
      }
      // Only the cell carrying ref owns the range. A shared dependent names the
      // si alone, and any of them may be written self closing.
      if (event.localName === 'f') {
        const kind = event.attributes.get('t')
        const covers = event.attributes.get('ref')
        if (covers !== undefined) {
          if (kind === 'shared') {
            const si = event.attributes.get('si')
            if (si !== undefined) master = { kind: 'shared', name: si }
          } else if (kind === 'array' || kind === 'dataTable') {
            master = { kind, name: covers }
          }
        }
      }
      continue
    }

    if (event.kind !== 'close') continue

    if (event.localName === 'c' && openCell) {
      currentCells.push({
        column: cellColumn,
        start: cellStart,
        end: event.end,
        style: cellStyle,
        spillingFormula: master,
      })
      openCell = false
      continue
    }
    if (event.localName === 'row' && currentRow !== undefined) {
      rows.push({
        ...currentRow,
        contentEnd: event.start,
        end: event.end,
        selfClosing: false,
        cells: currentCells,
      })
      currentRow = undefined
      continue
    }
    if (event.localName === 'dimension' && openDimension !== -1 && dimension !== undefined) {
      // Replacing only the open tag would leave the close tag behind.
      dimension = { start: openDimension, end: event.end, ref: dimension.ref }
      openDimension = -1
      continue
    }
    if (event.localName === 'sheetData') {
      contentEnd = event.start
      dataEnd = event.end
    }
    if (event.localName === 'autoFilter' && openAutoFilter !== -1) {
      autoFilter = { start: openAutoFilter, end: event.end }
      openAutoFilter = -1
    }
    if (event.localName === 'mergeCells') mergeInsertAt = event.start
    if (event.localName === 'cols') colInsertAt = event.start
    if (event.localName === 'dataValidations') dvInsertAt = event.start
    if (event.localName === 'worksheet') worksheetEnd = event.start
  }

  if (dataStart === -1)
    throw new XlsxError('malformed-xml', 'Sheet has no sheetData element to write into', at)

  return {
    prefix,
    dimension,
    rows,
    merges,
    contentEnd,
    selfClosing,
    dataStart,
    dataEnd,
    protection,
    autoFilter,
    sheetView,
    pane,
    worksheetStart,
    sheetPr,
    tabColor,
    sheetFormatPr,
    dataValidations:
      dvOpenStart === -1
        ? undefined
        : {
            openStart: dvOpenStart,
            openEnd: dvOpenEnd,
            insertAt: dvInsertAt,
            selfClosing: dvSelfClosing,
            count: dvCount,
          },
    laterSiblingStart,
    worksheetEnd,
    maxPriority,
    mergeContainer:
      mergeOpenStart === -1
        ? undefined
        : {
            openStart: mergeOpenStart,
            openEnd: mergeOpenEnd,
            insertAt: mergeInsertAt,
            selfClosing: mergeSelfClosing,
            count: mergeCount,
          },
    colContainer:
      colOpenStart === -1
        ? undefined
        : {
            openStart: colOpenStart,
            openEnd: colOpenEnd,
            insertAt: colInsertAt,
            selfClosing: colSelfClosing,
          },
    cols,
  }
}

interface ByteSplice {
  readonly start: number
  readonly end: number
  readonly text: string
  /** Orders insertions that land on the same offset. */
  readonly order: number
}

// The CT_Worksheet children that follow dataValidations. A new dataValidations is
// placed before the first of these the sheet has, so the schema order holds and
// Excel does not offer to repair the file; with none, it goes before the close.
const AFTER_DATA_VALIDATIONS = new Set([
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'drawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
  'legacyDrawing',
  'legacyDrawingHF',
])

/** One data-validation rule, built into a `<dataValidation>` by `patchSheet`. */
export interface DataValidationSpec {
  readonly type: string
  /** The range or cell the rule covers, canonical (`B2:B10`). */
  readonly sqref: string
  readonly allowBlank: boolean
  /** The comparison for a numeric rule; absent for a list. */
  readonly operator?: string
  /** Formula content, escaped when written. A list is a quoted, comma-joined set. */
  readonly formula1: string
  /** The upper bound of a two-sided comparison like `between`. */
  readonly formula2?: string
}

function dataValidationElement(spec: DataValidationSpec, prefix: string): string {
  const operator = spec.operator === undefined ? '' : ` operator="${spec.operator}"`
  const formula2 =
    spec.formula2 === undefined
      ? ''
      : `<${prefix}formula2>${escapeXml(spec.formula2)}</${prefix}formula2>`
  return (
    `<${prefix}dataValidation type="${spec.type}"${operator}` +
    ` allowBlank="${spec.allowBlank ? '1' : '0'}"` +
    ` showInputMessage="1" showErrorMessage="1" sqref="${spec.sqref}">` +
    `<${prefix}formula1>${escapeXml(spec.formula1)}</${prefix}formula1>${formula2}` +
    `</${prefix}dataValidation>`
  )
}

/** One conditional-format rule, built into a conditionalFormatting by `patchSheet`. */
export type ConditionalFormatSpec =
  | {
      readonly kind: 'colorScale'
      readonly sqref: string
      /** Two ARGB stops for a two-colour scale, three for a mid-pointed one. */
      readonly colors: readonly string[]
    }
  | {
      readonly kind: 'cellIs'
      readonly sqref: string
      readonly operator: string
      /** One formula for most comparisons, two for `between`/`notBetween`. */
      readonly formulas: readonly string[]
      /** Index of the dxf in styles.xml holding the highlight. */
      readonly dxfId: number
    }
  | { readonly kind: 'dataBar'; readonly sqref: string; readonly color: string }

function conditionalFormattingElement(
  spec: ConditionalFormatSpec,
  priority: number,
  prefix: string,
): string {
  const open = `<${prefix}conditionalFormatting sqref="${spec.sqref}">`
  const close = `</${prefix}conditionalFormatting>`
  if (spec.kind === 'colorScale') {
    const cfvo =
      spec.colors.length === 2
        ? `<${prefix}cfvo type="min"/><${prefix}cfvo type="max"/>`
        : `<${prefix}cfvo type="min"/><${prefix}cfvo type="percentile" val="50"/><${prefix}cfvo type="max"/>`
    const colors = spec.colors.map((color) => `<${prefix}color rgb="${color}"/>`).join('')
    return (
      `${open}<${prefix}cfRule type="colorScale" priority="${priority}">` +
      `<${prefix}colorScale>${cfvo}${colors}</${prefix}colorScale>` +
      `</${prefix}cfRule>${close}`
    )
  }
  if (spec.kind === 'dataBar') {
    return (
      `${open}<${prefix}cfRule type="dataBar" priority="${priority}"><${prefix}dataBar>` +
      `<${prefix}cfvo type="min"/><${prefix}cfvo type="max"/><${prefix}color rgb="${spec.color}"/>` +
      `</${prefix}dataBar></${prefix}cfRule>${close}`
    )
  }
  const formulas = spec.formulas
    .map((formula) => `<${prefix}formula>${escapeXml(formula)}</${prefix}formula>`)
    .join('')
  return (
    `${open}<${prefix}cfRule type="cellIs" operator="${spec.operator}" dxfId="${spec.dxfId}"` +
    ` priority="${priority}">${formulas}</${prefix}cfRule>${close}`
  )
}

/** Edits to a sheet that are not cell values: the elements around sheetData. */
export interface SheetEdits {
  readonly protection?: SheetProtection | 'remove'
  /** Canonical ranges (`A1:B2`) to merge, added to any the sheet already has. */
  readonly merges?: readonly string[]
  /** Row number to height in points, applied to the row's open tag. */
  readonly rowHeights?: ReadonlyMap<number, number>
  /** One-based column index to width, landing in the cols element before sheetData. */
  readonly columnWidths?: ReadonlyMap<number, number>
  /** Canonical range (`A1:B2`) for the sheet's autoFilter, replacing any it has. */
  readonly autoFilter?: string
  /** Canonical cell (`B2`) to freeze the panes above and left of, replacing any. */
  readonly freeze?: string
  /** One-based rows to hide, composed with any height the row also has. */
  readonly hiddenRows?: ReadonlySet<number>
  /** One-based columns to hide, composed with any width the column also has. */
  readonly hiddenColumns?: ReadonlySet<number>
  /** 8-digit ARGB for the sheet tab colour, into sheetPr, replacing any it has. */
  readonly tabColor?: string
  /** Whether the sheet shows gridlines, as `showGridLines` on the first sheetView. */
  readonly showGridLines?: boolean
  /** Whether the sheet shows row/column headings, as `showRowColHeaders`. */
  readonly showRowColHeaders?: boolean
  /** Zoom as a whole percentage, as `zoomScale` on the first sheetView. */
  readonly zoomScale?: number
  /** One-based row to its outline level, written as `outlineLevel` on the row. */
  readonly rowOutlineLevels?: ReadonlyMap<number, number>
  /** One-based column to its outline level, written as `outlineLevel` on the col. */
  readonly colOutlineLevels?: ReadonlyMap<number, number>
  /** Data-validation rules to add, joining any dataValidations the sheet has. */
  readonly dataValidations?: readonly DataValidationSpec[]
  /** Conditional-format rules to add, each its own conditionalFormatting element. */
  readonly conditionalFormats?: readonly ConditionalFormatSpec[]
}

export function patchSheet(
  bytes: Uint8Array,
  edits: ReadonlyMap<string, CellInput>,
  date1904: boolean,
  sharedStrings?: ReadonlyMap<string, number>,
  styleOverrides?: ReadonlyMap<string, number>,
  at: SheetLocation = {},
  sheet: SheetEdits = {},
): Uint8Array {
  const { protection, merges, autoFilter, freeze, tabColor } = sheet
  // A style override with no value edit is a restyle: the cell keeps its value
  // and formula and only its `s` changes, so unlike a write it needs no shared
  // formula or merge refusal — formatting is always safe.
  const restyleRefs =
    styleOverrides === undefined ? [] : [...styleOverrides.keys()].filter((ref) => !edits.has(ref))
  const newMerges = merges ?? []
  const rowHeights = sheet.rowHeights ?? new Map<number, number>()
  const columnWidths = sheet.columnWidths ?? new Map<number, number>()
  const hiddenRows = sheet.hiddenRows ?? new Set<number>()
  const hiddenColumns = sheet.hiddenColumns ?? new Set<number>()
  const rowOutlineLevels = sheet.rowOutlineLevels ?? new Map<number, number>()
  const colOutlineLevels = sheet.colOutlineLevels ?? new Map<number, number>()
  if (
    edits.size === 0 &&
    restyleRefs.length === 0 &&
    protection === undefined &&
    newMerges.length === 0 &&
    rowHeights.size === 0 &&
    columnWidths.size === 0 &&
    autoFilter === undefined &&
    freeze === undefined &&
    tabColor === undefined &&
    sheet.showGridLines === undefined &&
    sheet.showRowColHeaders === undefined &&
    sheet.zoomScale === undefined &&
    hiddenRows.size === 0 &&
    hiddenColumns.size === 0 &&
    rowOutlineLevels.size === 0 &&
    colOutlineLevels.size === 0 &&
    (sheet.dataValidations === undefined || sheet.dataValidations.length === 0) &&
    (sheet.conditionalFormats === undefined || sheet.conditionalFormats.length === 0)
  ) {
    return bytes
  }

  const rowAttributes = (row: number): string => {
    const height = rowHeights.get(row)
    const heightPart = height === undefined ? '' : ` ht="${height}" customHeight="1"`
    const level = rowOutlineLevels.get(row)
    const levelPart = level === undefined ? '' : ` outlineLevel="${level}"`
    return heightPart + (hiddenRows.has(row) ? ' hidden="1"' : '') + levelPart
  }
  const rowAttributed = (openTag: string, row: number): string => {
    const height = rowHeights.get(row)
    const withHeight =
      height === undefined
        ? openTag
        : withAttribute(withAttribute(openTag, 'ht', height), 'customHeight', 1)
    const withHidden = hiddenRows.has(row) ? withAttribute(withHeight, 'hidden', 1) : withHeight
    const level = rowOutlineLevels.get(row)
    return level === undefined ? withHidden : withAttribute(withHidden, 'outlineLevel', level)
  }

  const operations: ReadonlyArray<{ given: string; value: CellInput; restyle: boolean }> = [
    ...[...edits].map(([given, value]) => ({ given, value, restyle: false })),
    ...restyleRefs.map((given): { given: string; value: CellInput; restyle: boolean } => ({
      given,
      value: null,
      restyle: true,
    })),
  ]

  const shape = readShape(bytes, at)
  const splices: ByteSplice[] = []
  const newRows = new Map<number, string[]>()
  const filledRows = new Map<RowSpan, Array<{ column: number; cell: string }>>()

  const styleFor = (reference: string, current: string | undefined) => {
    const override = styleOverrides?.get(reference)
    return override === undefined ? current : String(override)
  }

  // Indexed rather than scanned, so writing many cells into a large sheet stays
  // linear in the number of edits instead of edits times rows.
  const rowsByNumber = new Map<number, RowSpan>()
  for (const candidate of shape.rows) rowsByNumber.set(candidate.row, candidate)

  const cellIndexes = new Map<RowSpan, Map<number, CellSpan>>()
  const cellsOf = (span: RowSpan) => {
    const known = cellIndexes.get(span)
    if (known !== undefined) return known
    const built = new Map<number, CellSpan>()
    for (const candidate of span.cells) built.set(candidate.column, candidate)
    cellIndexes.set(span, built)
    return built
  }

  for (const { given, value, restyle } of operations) {
    const address = parseReference(given)
    const { row, column } = address
    // The file never receives a reference spelled the way the caller typed it.
    const reference = formatReference(address)
    const existingRow = rowsByNumber.get(row)

    if (existingRow === undefined) {
      const pending = newRows.get(row) ?? []
      pending.push(
        cellElement(
          reference,
          value,
          styleFor(reference, undefined),
          date1904,
          sharedStrings,
          shape.prefix,
          at,
        ),
      )
      newRows.set(row, pending)
      continue
    }

    const existingCell = cellsOf(existingRow).get(column)
    if (existingCell !== undefined && restyle) {
      // Only the `s` on the open tag changes; the value and formula are kept.
      const override = styleOverrides?.get(reference)
      if (override !== undefined) {
        const element = decoder.decode(bytes.subarray(existingCell.start, existingCell.end))
        const openEnd = element.indexOf('>') + 1
        splices.push({
          start: existingCell.start,
          end: existingCell.end,
          text: withAttribute(element.slice(0, openEnd), 's', override) + element.slice(openEnd),
          order: column,
        })
      }
      continue
    }
    if (existingCell?.spillingFormula !== undefined) {
      throw sharedFormulaRefusal(reference, existingCell.spillingFormula, at)
    }
    if (existingCell !== undefined) {
      splices.push({
        start: existingCell.start,
        end: existingCell.end,
        text: cellElement(
          reference,
          value,
          styleFor(reference, existingCell.style),
          date1904,
          sharedStrings,
          shape.prefix,
          at,
        ),
        order: column,
      })
      continue
    }

    const cell = cellElement(
      reference,
      value,
      styleFor(reference, undefined),
      date1904,
      sharedStrings,
      shape.prefix,
      at,
    )

    if (existingRow.selfClosing) {
      // Collected rather than spliced now: two cells added to the same empty
      // row would otherwise each rewrite it, emitting the row twice.
      const pending = filledRows.get(existingRow) ?? []
      pending.push({ column, cell })
      filledRows.set(existingRow, pending)
      continue
    }

    const next = existingRow.cells.find((candidate) => candidate.column > column)
    const insertAt = next === undefined ? existingRow.contentEnd : next.start
    splices.push({ start: insertAt, end: insertAt, text: cell, order: column })
  }

  for (const [row, cells] of filledRows) {
    const ordered = [...cells]
      .sort((left, right) => left.column - right.column)
      .map((entry) => entry.cell)
      .join('')
    const openTag = decoder.decode(bytes.subarray(row.start, row.end))
    const opened = rowAttributed(openTag, row.row)
    splices.push({
      start: row.start,
      end: row.end,
      text: `${opened.slice(0, -2)}>${ordered}</${shape.prefix}row>`,
      order: row.row,
    })
  }

  const editedRows = new Set([...rowHeights.keys(), ...hiddenRows, ...rowOutlineLevels.keys()])

  // A height or a hide on a row that neither exists nor has a value edit is a new
  // empty row.
  for (const row of editedRows) {
    if (!rowsByNumber.has(row) && !newRows.has(row)) newRows.set(row, [])
  }

  // A row that stays as it is has its open tag rewritten, unless a cell was added
  // to a self closing one, where filledRows already redid the whole row.
  for (const row of editedRows) {
    const span = rowsByNumber.get(row)
    if (span === undefined || filledRows.has(span)) continue
    const openTag = decoder.decode(bytes.subarray(span.start, span.openEnd))
    splices.push({
      start: span.start,
      end: span.openEnd,
      text: rowAttributed(openTag, row),
      order: span.row,
    })
  }

  const buildRow = (row: number, cells: string[]) => {
    const ordered = cells
      .map((text) => ({ text, column: parseReference(cellReferenceOf(text)).column }))
      .sort((a, b) => a.column - b.column)
      .map((entry) => entry.text)
      .join('')
    return `<${shape.prefix}row r="${row}"${rowAttributes(row)}>${ordered}</${shape.prefix}row>`
  }

  if (shape.selfClosing && newRows.size > 0) {
    const body = [...newRows]
      .sort(([left], [right]) => left - right)
      .map(([row, cells]) => buildRow(row, cells))
      .join('')
    splices.push({
      start: shape.dataStart,
      end: shape.dataEnd,
      text: `<${shape.prefix}sheetData>${body}</${shape.prefix}sheetData>`,
      order: 0,
    })
  } else {
    // Existing rows come out of readShape in document order, so walking them
    // once alongside the sorted new rows places every one without rescanning
    // the sheet per row.
    let cursor = 0
    let next = shape.rows[cursor]
    for (const [row, cells] of [...newRows].sort(([left], [right]) => left - right)) {
      while (next !== undefined && next.row <= row) {
        cursor++
        next = shape.rows[cursor]
      }
      const offset = next === undefined ? shape.contentEnd : next.start
      splices.push({ start: offset, end: offset, text: buildRow(row, cells), order: row })
    }
  }

  const widened = widenDimension(shape.dimension, [...edits.keys()])
  if (widened !== undefined && shape.dimension !== undefined) {
    splices.push({
      start: shape.dimension.start,
      end: shape.dimension.end,
      text: `<${shape.prefix}dimension ref="${widened}"/>`,
      order: -1,
    })
  }

  // sheetProtection is the sibling after sheetData, so it goes at dataEnd unless
  // the sheet already has one to replace or remove in place.
  if (protection === 'remove') {
    if (shape.protection !== undefined) {
      splices.push({
        start: shape.protection.start,
        end: shape.protection.end,
        text: '',
        order: -1,
      })
    }
  } else if (protection !== undefined) {
    const span = shape.protection ?? { start: shape.dataEnd, end: shape.dataEnd }
    splices.push({
      start: span.start,
      end: span.end,
      text: sheetProtectionElement(protection, shape.prefix),
      order: -1,
    })
  }

  // autoFilter sits between sheetProtection and mergeCells. It replaces the one
  // the sheet has, or lands just after sheetData/sheetProtection, ordered before
  // mergeCells at a shared offset.
  if (autoFilter !== undefined) {
    const anchor = shape.protection?.end ?? shape.dataEnd
    const span = shape.autoFilter ?? { start: anchor, end: anchor }
    splices.push({
      start: span.start,
      end: span.end,
      text: `<${shape.prefix}autoFilter ref="${autoFilter}"/>`,
      order: -0.5,
    })
  }

  // The first sheetView, before cols and sheetData, carries a freeze pane and the
  // view flags (gridlines, headings, zoom). They share one element, so both are
  // handled together: attributes rewrite its open tag, a pane goes in as its first
  // child. Modifying an existing sheetView leaves any other children it has alone.
  const viewAttributes: Array<[string, number]> = []
  if (sheet.showGridLines !== undefined)
    viewAttributes.push(['showGridLines', sheet.showGridLines ? 1 : 0])
  if (sheet.showRowColHeaders !== undefined)
    viewAttributes.push(['showRowColHeaders', sheet.showRowColHeaders ? 1 : 0])
  if (sheet.zoomScale !== undefined) viewAttributes.push(['zoomScale', sheet.zoomScale])

  if (freeze !== undefined || viewAttributes.length > 0) {
    let paneXml: string | undefined
    if (freeze !== undefined) {
      const { column, row } = parseReference(freeze)
      const xSplit = column - 1
      const ySplit = row - 1
      const activePane =
        xSplit > 0 && ySplit > 0 ? 'bottomRight' : xSplit > 0 ? 'topRight' : 'bottomLeft'
      paneXml =
        `<${shape.prefix}pane` +
        (xSplit > 0 ? ` xSplit="${xSplit}"` : '') +
        (ySplit > 0 ? ` ySplit="${ySplit}"` : '') +
        ` topLeftCell="${freeze}" activePane="${activePane}" state="frozen"/>`
    }

    const withViewAttributes = (openTag: string): string =>
      viewAttributes.reduce((tag, [name, value]) => withAttribute(tag, name, value), openTag)

    const view = shape.sheetView
    if (view === undefined) {
      const attrs = viewAttributes.map(([name, value]) => ` ${name}="${value}"`).join('')
      const open = `<${shape.prefix}sheetView workbookViewId="0"${attrs}/>`
      const element =
        paneXml === undefined ? open : `${open.slice(0, -2)}>${paneXml}</${shape.prefix}sheetView>`
      const anchor = shape.colContainer?.openStart ?? shape.dataStart
      splices.push({
        start: anchor,
        end: anchor,
        text: `<${shape.prefix}sheetViews>${element}</${shape.prefix}sheetViews>`,
        order: -3,
      })
    } else {
      const openTag = decoder.decode(bytes.subarray(view.start, view.end))
      if (shape.pane !== undefined) {
        if (paneXml !== undefined)
          splices.push({ start: shape.pane.start, end: shape.pane.end, text: paneXml, order: 0 })
        if (viewAttributes.length > 0)
          splices.push({
            start: view.start,
            end: view.end,
            text: withViewAttributes(openTag),
            order: -1,
          })
      } else if (view.selfClosing) {
        const open = withViewAttributes(openTag)
        const text =
          paneXml === undefined
            ? open
            : `${open.slice(0, -2)}>${paneXml}</${shape.prefix}sheetView>`
        splices.push({ start: view.start, end: view.end, text, order: 0 })
      } else {
        if (viewAttributes.length > 0)
          splices.push({
            start: view.start,
            end: view.end,
            text: withViewAttributes(openTag),
            order: -1,
          })
        if (paneXml !== undefined)
          splices.push({ start: view.end, end: view.end, text: paneXml, order: -1 })
      }
    }
  }

  // tabColor is the first child of sheetPr, itself the first child of worksheet.
  // It replaces the tabColor the sheet has, becomes the first child of an existing
  // sheetPr, or brings a fresh sheetPr at the very front.
  if (tabColor !== undefined) {
    const tabXml = `<${shape.prefix}tabColor rgb="${tabColor}"/>`
    if (shape.tabColor !== undefined) {
      splices.push({ start: shape.tabColor.start, end: shape.tabColor.end, text: tabXml, order: 0 })
    } else if (shape.sheetPr !== undefined) {
      if (shape.sheetPr.selfClosing) {
        const openTag = decoder.decode(bytes.subarray(shape.sheetPr.start, shape.sheetPr.end))
        splices.push({
          start: shape.sheetPr.start,
          end: shape.sheetPr.end,
          text: `${openTag.slice(0, -2)}>${tabXml}</${shape.prefix}sheetPr>`,
          order: 0,
        })
      } else {
        splices.push({ start: shape.sheetPr.end, end: shape.sheetPr.end, text: tabXml, order: -1 })
      }
    } else {
      splices.push({
        start: shape.worksheetStart,
        end: shape.worksheetStart,
        text: `<${shape.prefix}sheetPr>${tabXml}</${shape.prefix}sheetPr>`,
        order: -5,
      })
    }
  }

  // mergeCells is the sibling after sheetProtection. New ranges join the sheet's
  // own mergeCells when it has one, or open a fresh one past sheetData.
  if (newMerges.length > 0) {
    const seen = new Set(shape.merges.map(canonicalMerge))
    const fresh: string[] = []
    for (const range of newMerges) {
      if (seen.has(range)) continue
      seen.add(range)
      fresh.push(range)
    }
    if (fresh.length > 0) {
      const children = fresh.map((ref) => `<${shape.prefix}mergeCell ref="${ref}"/>`).join('')
      const container = shape.mergeContainer
      if (container === undefined) {
        const anchor = shape.protection?.end ?? shape.dataEnd
        splices.push({
          start: anchor,
          end: anchor,
          text: `<${shape.prefix}mergeCells count="${fresh.length}">${children}</${shape.prefix}mergeCells>`,
          order: 0,
        })
      } else {
        const openTag = decoder.decode(bytes.subarray(container.openStart, container.openEnd))
        const counted = withAttribute(openTag, 'count', container.count + fresh.length)
        if (container.selfClosing) {
          splices.push({
            start: container.openStart,
            end: container.openEnd,
            text: `${counted.slice(0, -2)}>${children}</${shape.prefix}mergeCells>`,
            order: -1,
          })
        } else {
          splices.push({
            start: container.openStart,
            end: container.openEnd,
            text: counted,
            order: -1,
          })
          splices.push({
            start: container.insertAt,
            end: container.insertAt,
            text: children,
            order: 0,
          })
        }
      }
    }
  }

  // cols is the sibling before sheetData. A width or a hide lands in the col whose
  // range covers its column, splitting that range when it spans more than the
  // column, or opening a new col when none covers it.
  if (columnWidths.size > 0 || hiddenColumns.size > 0 || colOutlineLevels.size > 0) {
    const colAttributes = (column: number): string => {
      const width = columnWidths.get(column)
      const widthPart = width === undefined ? '' : ` width="${width}" customWidth="1"`
      const level = colOutlineLevels.get(column)
      const levelPart = level === undefined ? '' : ` outlineLevel="${level}"`
      return widthPart + (hiddenColumns.has(column) ? ' hidden="1"' : '') + levelPart
    }
    const colElement = (column: number): string =>
      `<${shape.prefix}col min="${column}" max="${column}"${colAttributes(column)}/>`
    const colWith = (element: string, min: number, max: number, edited?: number): string => {
      const ranged = withAttribute(withAttribute(element, 'min', min), 'max', max)
      if (edited === undefined) return ranged
      const width = columnWidths.get(edited)
      const withWidth =
        width === undefined
          ? ranged
          : withAttribute(withAttribute(ranged, 'width', width), 'customWidth', 1)
      const withHidden = hiddenColumns.has(edited)
        ? withAttribute(withWidth, 'hidden', 1)
        : withWidth
      const level = colOutlineLevels.get(edited)
      return level === undefined ? withHidden : withAttribute(withHidden, 'outlineLevel', level)
    }

    const covering = (column: number) => shape.cols.find((c) => c.min <= column && column <= c.max)
    const bySpan = new Map<(typeof shape.cols)[number], number[]>()
    const appends: number[] = []
    for (const column of new Set([
      ...columnWidths.keys(),
      ...hiddenColumns,
      ...colOutlineLevels.keys(),
    ])) {
      const span = covering(column)
      if (span === undefined) {
        appends.push(column)
        continue
      }
      const grouped = bySpan.get(span) ?? []
      grouped.push(column)
      bySpan.set(span, grouped)
    }

    for (const [span, columns] of bySpan) {
      const element = decoder.decode(bytes.subarray(span.start, span.end))
      const segments: string[] = []
      let cursor = span.min
      for (const column of [...columns].sort((a, b) => a - b)) {
        if (column > cursor) segments.push(colWith(element, cursor, column - 1))
        segments.push(colWith(element, column, column, column))
        cursor = column + 1
      }
      if (cursor <= span.max) segments.push(colWith(element, cursor, span.max))
      splices.push({ start: span.start, end: span.end, text: segments.join(''), order: span.min })
    }

    if (appends.length > 0) {
      const added = [...appends]
        .sort((a, b) => a - b)
        .map((column) => colElement(column))
        .join('')
      const container = shape.colContainer
      if (container === undefined) {
        splices.push({
          start: shape.dataStart,
          end: shape.dataStart,
          text: `<${shape.prefix}cols>${added}</${shape.prefix}cols>`,
          order: -2,
        })
      } else if (container.selfClosing) {
        const openTag = decoder.decode(bytes.subarray(container.openStart, container.openEnd))
        splices.push({
          start: container.openStart,
          end: container.openEnd,
          text: `${openTag.slice(0, -2)}>${added}</${shape.prefix}cols>`,
          order: -1,
        })
      } else {
        splices.push({ start: container.insertAt, end: container.insertAt, text: added, order: 1 })
      }
    }
  }

  // sheetFormatPr's outlineLevelRow/Col hint the deepest group. It is updated when
  // the sheet has a sheetFormatPr, taking the max of what it declared and what was
  // grouped; a sheet without one still groups by the outlineLevel on its rows and
  // cols, so no sheetFormatPr is invented (it would need a defaultRowHeight).
  if (
    (rowOutlineLevels.size > 0 || colOutlineLevels.size > 0) &&
    shape.sheetFormatPr !== undefined
  ) {
    const format = decoder.decode(
      bytes.subarray(shape.sheetFormatPr.start, shape.sheetFormatPr.end),
    )
    const declared = (attribute: string): number => {
      const match = format.match(new RegExp(`\\b${attribute}="(\\d+)"`))
      return match === null ? 0 : Number(match[1])
    }
    const maxRow = Math.max(declared('outlineLevelRow'), ...rowOutlineLevels.values())
    const maxCol = Math.max(declared('outlineLevelCol'), ...colOutlineLevels.values())
    let tag = format
    if (maxRow > 0) tag = withAttribute(tag, 'outlineLevelRow', maxRow)
    if (maxCol > 0) tag = withAttribute(tag, 'outlineLevelCol', maxCol)
    splices.push({
      start: shape.sheetFormatPr.start,
      end: shape.sheetFormatPr.end,
      text: tag,
      order: 0,
    })
  }

  // conditionalFormatting sits just before dataValidations. Each rule is its own
  // element, added after any the sheet has, before dataValidations or the first
  // later sibling, or before the worksheet close. Its priority outranks the
  // highest the sheet already uses, so a new rule wins ties of evaluation order.
  const conditionalFormats = sheet.conditionalFormats ?? []
  if (conditionalFormats.length > 0) {
    const candidates: number[] = []
    if (shape.dataValidations !== undefined) candidates.push(shape.dataValidations.openStart)
    if (shape.laterSiblingStart !== -1) candidates.push(shape.laterSiblingStart)
    const anchor = candidates.length === 0 ? shape.worksheetEnd : Math.min(...candidates)
    const elements = conditionalFormats
      .map((spec, index) =>
        conditionalFormattingElement(spec, shape.maxPriority + 1 + index, shape.prefix),
      )
      .join('')
    splices.push({ start: anchor, end: anchor, text: elements, order: -0.7 })
  }

  // dataValidations sits after mergeCells and conditionalFormatting, before
  // hyperlinks and the page-setup family. New rules join the sheet's own
  // dataValidations when it has one, or open a fresh one at the schema-correct
  // spot: before the first later sibling, or before the worksheet close.
  const validations = sheet.dataValidations ?? []
  if (validations.length > 0) {
    const elements = validations.map((spec) => dataValidationElement(spec, shape.prefix)).join('')
    const existing = shape.dataValidations
    if (existing === undefined) {
      const anchor = shape.laterSiblingStart === -1 ? shape.worksheetEnd : shape.laterSiblingStart
      splices.push({
        start: anchor,
        end: anchor,
        text: `<${shape.prefix}dataValidations count="${validations.length}">${elements}</${shape.prefix}dataValidations>`,
        order: -0.5,
      })
    } else {
      const openTag = decoder.decode(bytes.subarray(existing.openStart, existing.openEnd))
      const counted = withAttribute(openTag, 'count', existing.count + validations.length)
      if (existing.selfClosing) {
        splices.push({
          start: existing.openStart,
          end: existing.openEnd,
          text: `${counted.slice(0, -2)}>${elements}</${shape.prefix}dataValidations>`,
          order: 0,
        })
      } else {
        splices.push({ start: existing.openStart, end: existing.openEnd, text: counted, order: -1 })
        splices.push({ start: existing.insertAt, end: existing.insertAt, text: elements, order: 0 })
      }
    }
  }

  return applyByteSplices(bytes, splices)
}

/**
 * Excel recalculates the used range, but stricter readers trust what the file
 * declares, so a cell written outside it would be ignored.
 */
function widenDimension(
  dimension: DimensionSpan | undefined,
  references: readonly string[],
): string | undefined {
  if (dimension === undefined) return undefined

  const bounds = dimension.ref.split(':')
  const from = bounds[0] ?? ''
  if (from === '') return undefined

  // The dimension comes from the file, so its halves may not parse, and a range
  // it names past the last column cannot be spelled back. Either way the file's
  // dimension is left as it is rather than taking the whole save down; the cell
  // still writes, and Excel recalculates the used range anyway.
  try {
    const topLeft = parseReference(from)
    const bottomRight = parseReference(bounds[1] ?? from)

    let { row: lastRow, column: lastColumn } = bottomRight
    let { row: firstRow, column: firstColumn } = topLeft

    for (const reference of references) {
      const { row, column } = parseReference(reference)
      firstRow = Math.min(firstRow, row)
      firstColumn = Math.min(firstColumn, column)
      lastRow = Math.max(lastRow, row)
      lastColumn = Math.max(lastColumn, column)
    }

    const grown =
      firstRow !== topLeft.row ||
      firstColumn !== topLeft.column ||
      lastRow !== bottomRight.row ||
      lastColumn !== bottomRight.column
    if (!grown) return undefined

    const start = formatReference({ row: firstRow, column: firstColumn })
    const end = formatReference({ row: lastRow, column: lastColumn })
    return `${start}:${end}`
  } catch {
    return undefined
  }
}

/** Only safe on elements built above, where the reference is the first attribute. */
function cellReferenceOf(element: string): string {
  const start = element.indexOf('"') + 1
  return element.slice(start, element.indexOf('"', start))
}

const decoder = new TextDecoder()
const encoder = new TextEncoder()

function applyByteSplices(bytes: Uint8Array, splices: ByteSplice[]): Uint8Array {
  const ordered = [...splices].sort((a, b) => a.start - b.start || a.order - b.order)

  const pieces: Uint8Array[] = []
  let cursor = 0
  for (const splice of ordered) {
    pieces.push(bytes.subarray(cursor, splice.start))
    pieces.push(encoder.encode(splice.text))
    cursor = splice.end
  }
  pieces.push(bytes.subarray(cursor))

  let length = 0
  for (const piece of pieces) length += piece.length
  const out = new Uint8Array(length)
  let offset = 0
  for (const piece of pieces) {
    out.set(piece, offset)
    offset += piece.length
  }
  return out
}
