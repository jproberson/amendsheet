import { type Container, decodeXmlPart, writeContainer } from './container.js'
import { XlsxError } from './errors.js'
import { LAST_SERIAL, dateToSerial, serialToDate } from './date.js'
import {
  type CellInput,
  checkWritable,
  mergeAnchorFor,
  mergeRefusal,
  patchSheet,
  indexSheet,
  sharedFormulaRefusal,
  type SheetIndex,
  type SheetLocation,
} from './patch.js'
import {
  type CellAddress,
  canonicalReference,
  formatReference,
  parseReference,
  parseWritableReference,
} from './reference.js'
import { resolveTarget } from './relationships.js'
import { type RawCell, readSheet } from './sheet.js'
import { appendSharedStrings, readSharedStrings } from './shared-strings.js'
import { extendTables } from './tables.js'
import { type DateStyle, ensureDateStyle, ensureNumberFormat } from './styles-writer.js'
import { type Styles, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { readXml } from './xml.js'
import { type SheetState, readWorkbookPart } from './workbook.js'

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
}

export interface Worksheet {
  readonly name: string
  readonly state: SheetState
  /**
   * As the workbook part spells it, so a defined name or a part this library
   * does not interpret can be matched against the sheet it refers to.
   */
  readonly sheetId: string
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
}

export interface SetOptions {
  /** A number format code, applied to the cell being written. */
  readonly numberFormat?: string
}

export interface Workbook {
  readonly sheets: readonly Worksheet[]
  /** Undefined when no sheet has that name. Names are compared exactly. */
  sheet(name: string): Worksheet | undefined
  /** Which year serials count from. A 1904 workbook is 1462 days behind. */
  readonly epoch: 1900 | 1904
  /** Parts that were never interpreted are written exactly as they were read. */
  toBytes(): Uint8Array
}

const EMPTY_STYLES: Styles = { numberFormats: new Map(), cellFormats: [] }

const CALCULATION_CHAIN = 'xl/calcChain.xml'
const CONTENT_TYPES = '[Content_Types].xml'

/**
 * CT_Workbook is a sequence, so a calcPr the file lacks cannot simply be
 * appended: these are the children the schema puts after it.
 */
const AFTER_CALC_PR = new Set([
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
])

/**
 * A written formula carries no computed result, so the workbook is marked for
 * recalculation. Without it a reader that trusts cached values shows nothing.
 */
function withRecalculation(xml: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'calcPr') continue

    const tag = xml.slice(event.start, event.end)
    if (tag.includes('fullCalcOnLoad=')) {
      return (
        xml.slice(0, event.start) +
        // Either quote character is legal, and matching only double quotes
        // left the check passing while the rewrite did nothing.
        tag.replace(/fullCalcOnLoad=("|')[^"']*\1/, 'fullCalcOnLoad="1"') +
        xml.slice(event.end)
      )
    }
    const opened = tag.replace(/\/?>$/, (end) =>
      end === '/>' ? ' fullCalcOnLoad="1"/>' : ' fullCalcOnLoad="1">',
    )
    return xml.slice(0, event.start) + opened + xml.slice(event.end)
  }

  let depth = 0
  let insertAt = -1

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      // Direct children of the root open at depth 1, so a nested extLst in a
      // part we do not interpret cannot be mistaken for the workbook's own.
      if (depth === 1 && insertAt === -1 && AFTER_CALC_PR.has(event.localName)) {
        insertAt = event.start
      }
      if (!event.selfClosing) depth++
      continue
    }
    if (event.kind !== 'close') continue

    depth--
    if (event.localName !== 'workbook' || depth !== 0) continue

    const colon = event.name.indexOf(':')
    const prefix = colon === -1 ? '' : event.name.slice(0, colon + 1)
    const element = `<${prefix}calcPr fullCalcOnLoad="1"/>`
    const at = insertAt === -1 ? event.start : insertAt
    return xml.slice(0, at) + element + xml.slice(at)
  }

  return xml
}

/**
 * Removes the relationship pointing at one part, leaving every other byte
 * alone. A relationship whose target is gone is an invalid package.
 */
function withoutRelationship(xml: string, ownerPath: string, part: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'Relationship') continue
    if (event.attributes.get('TargetMode') === 'External') continue
    const target = event.attributes.get('Target')
    if (target === undefined || resolveTarget(ownerPath, target) !== part) continue
    return xml.slice(0, event.start) + xml.slice(event.end)
  }
  return xml
}

/** Removes one Override element, leaving every other byte of the part alone. */
function withoutOverride(xml: string, part: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'Override') continue
    if (event.attributes.get('PartName') !== `/${part}`) continue
    return xml.slice(0, event.start) + xml.slice(event.end)
  }
  return xml
}

function partText(container: Container, path: string): string | undefined {
  const bytes = container.parts.get(path)
  if (bytes === undefined) return undefined
  return decodeXmlPart(bytes, path)
}

function toCellValue(raw: RawCell, styles: Styles, date1904: boolean): CellValue {
  const value = raw.value

  if (value.kind === 'date') {
    const parsed = new Date(value.value)
    if (Number.isNaN(parsed.getTime())) return { kind: 'text', value: value.value }
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

function toCell(raw: RawCell, styles: Styles, date1904: boolean, masters?: SharedMasters): Cell {
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
  }
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

  const sheets = part.sheets.map((reference): Worksheet => {
    const sheetXml = partText(container, reference.path)
    const at: SheetLocation = { sheet: reference.name, part: reference.path }

    const patched = () => {
      if (sheetXml === undefined) return undefined
      const pending = edits.get(reference.path)
      if (pending === undefined) return sheetXml
      return patchSheet(
        sheetXml,
        pending,
        date1904,
        undefined,
        styleOverrides.get(reference.path),
        at,
      )
    }

    function* readCells(source?: string): Generator<Cell> {
      const xml = source ?? patched()
      if (xml === undefined) return

      // The cell owning a shared range comes before its dependents, so filling
      // this in as the sheet streams resolves every one of them.
      const masters: SharedMasters = new Map()

      for (const raw of readSheet(xml, sharedStrings, at)) {
        if (raw.ownsSharedRange === true && raw.sharedIndex !== undefined) {
          masters.set(raw.sharedIndex, canonicalReference(raw.address) ?? raw.reference)
        }
        yield toCell(raw, stylesNow(), date1904, masters)
      }
    }

    // The sheet as it was read: which style each cell carried, and which cells
    // define a shared formula. Built once, on the first write, because every
    // set() needs both and parsing the sheet twice per sheet showed up in the
    // bench.
    let index: SheetIndex | undefined
    const indexed = (): SheetIndex | undefined => {
      if (sheetXml === undefined) return undefined
      index ??= indexSheet(sheetXml)
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

      if (value === null) {
        return toCell({ ...raw, value: { kind: 'empty' } }, stylesNow(), date1904)
      }
      if (typeof value === 'number') {
        return toCell({ ...raw, value: { kind: 'number', value } }, stylesNow(), date1904)
      }
      if (typeof value === 'boolean') {
        return toCell({ ...raw, value: { kind: 'boolean', value } }, stylesNow(), date1904)
      }
      if (typeof value === 'string') {
        return toCell({ ...raw, value: { kind: 'text', value } }, stylesNow(), date1904)
      }
      if (value instanceof Date) {
        // Written as the serial it becomes, so the style decides how it reads.
        const serial = dateToSerial(value, date1904)
        return toCell({ ...raw, value: { kind: 'number', value: serial } }, stylesNow(), date1904)
      }
      // A formula is written without a cached value, so it reads back empty.
      const formula = value.formula
      return toCell({ ...raw, value: { kind: 'empty' }, formula }, stylesNow(), date1904)
    }

    // Built from the sheet as it was read. Every edit is in the overlay, so the
    // index never needs rebuilding.
    let byReference: Map<string, Cell> | undefined
    const overlay = new Map<string, Cell>()

    return {
      name: reference.name,
      state: reference.state,
      sheetId: reference.sheetId,
      cells: () => readCells(),
      cell(cellReference: string): Cell | undefined {
        const wanted = canonicalReference(parseReference(cellReference))
        if (wanted === undefined) return undefined
        const edited = overlay.get(wanted)
        if (edited !== undefined) return edited

        if (byReference === undefined) {
          byReference = new Map()
          for (const found of readCells(sheetXml)) {
            const at = canonicalReference(found.address)
            if (at !== undefined) byReference.set(at, found)
          }
        }
        return byReference.get(wanted)
      },
      set(cellReference: string, value: CellInput, options?: SetOptions): void {
        // Normalised so `a1`, `$A$1` and `A1` are one edit, and so the file
        // never receives a reference spelled the way the caller typed it.
        const canonical = formatReference(parseWritableReference(cellReference))

        // Refused here rather than at save time. An edit that only fails once
        // the workbook is written takes the whole batch down with it, and until
        // then cell() reports a write that is never going to happen.
        if (sheetXml === undefined) {
          // Recording it would report the value from cell() and save none of it.
          throw new XlsxError(
            'missing-part',
            `Sheet ${reference.name} is not in the package, so ${canonical} cannot be written`,
            { ...at, reference: canonical },
          )
        }

        checkWritable(canonical, value, date1904, at)
        // sheetXml is present, so indexed() is too; the guard is for the type.
        const index = indexed()
        if (index !== undefined) {
          const si = index.sharedFormulas.get(canonical)
          if (si !== undefined) throw sharedFormulaRefusal(canonical, si, at)
          const anchor = mergeAnchorFor(index, canonical)
          if (anchor !== undefined) throw mergeRefusal(canonical, anchor, at)
        }

        const current = styleAt(canonical)

        // Resolved before anything is recorded. A format this file cannot hold
        // refuses here, and a refusal that had already queued the edit would
        // write the value it claimed to reject.
        let applied: DateStyle | undefined
        if (workingStyles === undefined) {
          if (options?.numberFormat !== undefined) {
            throw new XlsxError(
              'missing-part',
              `Cannot apply a number format to ${canonical}: the package has no style table`,
              { part: 'xl/styles.xml', reference: canonical },
            )
          }
        } else {
          // An asked-for format wins; a Date only gets one because without one
          // it displays as the serial number it is stored as.
          if (options?.numberFormat !== undefined) {
            applied = ensureNumberFormat(workingStyles, current, options.numberFormat)
          } else if (value instanceof Date) {
            applied = ensureDateStyle(workingStyles, current)
          }
        }

        const pending = edits.get(reference.path) ?? new Map<string, CellInput>()
        pending.set(canonical, value)
        edits.set(reference.path, pending)

        let resolved = current
        if (applied !== undefined) {
          workingStyles = applied.xml
          resolved = applied.index
          if (applied.index !== current) {
            const overrides = styleOverrides.get(reference.path) ?? new Map<string, number>()
            overrides.set(canonical, applied.index)
            styleOverrides.set(reference.path, overrides)
          }
        }

        overlay.set(canonical, predict(canonical, value, resolved))
      },
    }
  })

  const toBytes = (): Uint8Array => {
    if (edits.size === 0) return writeContainer(container)

    const parts = new Map(container.parts)
    const encoder = new TextEncoder()

    // Excel rebuilds the calculation chain, but a stale one makes it offer to
    // repair the file, so the part and its content type go together.
    if (parts.delete(CALCULATION_CHAIN)) {
      const types = partText(container, CONTENT_TYPES)
      if (types !== undefined) {
        parts.set(CONTENT_TYPES, encoder.encode(withoutOverride(types, CALCULATION_CHAIN)))
      }
      const rels = partText(container, part.relationshipsPath)
      if (rels !== undefined) {
        parts.set(
          part.relationshipsPath,
          encoder.encode(withoutRelationship(rels, part.path, CALCULATION_CHAIN)),
        )
      }
    }

    // set() already resolved every style; only the serialising is left.
    if (workingStyles !== undefined && workingStyles !== stylesXml) {
      parts.set('xl/styles.xml', encoder.encode(workingStyles))
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
        parts.set('xl/sharedStrings.xml', encoder.encode(appended.xml))
        indexes = appended.indexes
      }
    }

    for (const [path, pending] of edits) {
      const xml = partText(container, path)
      if (xml === undefined) continue
      const at: SheetLocation = {
        sheet: part.sheets.find((s) => s.path === path)?.name,
        part: path,
      }
      parts.set(
        path,
        encoder.encode(patchSheet(xml, pending, date1904, indexes, styleOverrides.get(path), at)),
      )
      // A cell written just past a table grows it, the way Excel would.
      for (const extension of extendTables(xml, path, container, pending.keys())) {
        parts.set(extension.path, encoder.encode(extension.xml))
      }
    }

    const wroteFormula = [...edits.values()].some((pending) =>
      [...pending.values()].some(
        (value) => typeof value === 'object' && value !== null && !(value instanceof Date),
      ),
    )
    if (wroteFormula) {
      const book = partText(container, part.path)
      if (book !== undefined) parts.set(part.path, encoder.encode(withRecalculation(book)))
    }

    return writeContainer({ parts })
  }

  return {
    sheets,
    sheet: (name: string) => sheets.find((candidate) => candidate.name === name),
    epoch: date1904 ? 1904 : 1900,
    toBytes,
  }
}
