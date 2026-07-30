import type { XlsxErrorContext } from './errors.js'
import { XlsxError } from './errors.js'
import {
  type Alignment,
  type AlignmentAttributes,
  type BorderFormat,
  type CellFormatting,
  type CellProtection,
  type CellOverrides,
  DEFAULT_XF,
  type FillFormat,
  type FontFormat,
  FIRST_CUSTOM_FORMAT_ID,
  type ProtectionAttributes,
  RESERVED_BORDER,
  RESERVED_FILLS,
  SHORT_DATE_FORMAT_ID,
  SIDE_NAMES,
  type Side,
  type Sides,
  attrId,
  buildAlignmentElement,
  buildBorderElement,
  buildFillElement,
  buildFontElement,
  buildProtectionElement,
  escapeXml,
  formattingOfXf,
  isPlainFormat,
  parseAlignment,
  parseBorder,
  parseFont,
  parseProtection,
  tablePrefix as tablePrefixOf,
  usedFormatIds,
  validateAlignment,
  withAlignmentChild,
  withNamespacePrefix,
  withOverrides,
  withProtectionChild,
} from './styles-writer.js'
import { type Styles, builtInFormatId, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { findUnwritableCharacter, readXml, withAttribute } from './xml.js'

/**
 * A styles.xml amended in memory instead of by re-parsing and re-splicing the
 * growing string on every edit. The original is parsed once into per-table
 * element arrays and dedup maps; each write assigns an index against those in O(1)
 * amortised, and `serialize` splices all additions into the original once. The
 * result is byte-identical to threading the same edits through the `ensure*`
 * functions in `styles-writer.ts`, which stay the oracle the differential test
 * checks against.
 */
export interface StylesSession {
  numberFormat(basedOn: number | undefined, code: string, location: XlsxErrorContext): number
  dateStyle(basedOn: number | undefined): number
  font(basedOn: number | undefined, font: FontFormat, location: XlsxErrorContext): number
  fill(basedOn: number | undefined, fill: FillFormat, location: XlsxErrorContext): number
  border(basedOn: number | undefined, border: BorderFormat, location: XlsxErrorContext): number
  alignment(basedOn: number | undefined, alignment: Alignment, location: XlsxErrorContext): number
  protection(basedOn: number | undefined, protection: CellProtection): number
  /** Runs `body` so a throw part-way rolls every in-memory change back, matching
   * the oracle where a rejected edit leaves the styles string untouched. */
  transaction<T>(body: () => T): T
  /** The number format code a cell format resolves to, served from memory. */
  numberFormatOf(index: number): string | undefined
  /** The font/fill/border/alignment/protection a cell format carries, served from
   * memory so a read after a write does not force a re-serialize. */
  formattingOf(index: number): CellFormatting
  /** The number-format tables a read resolves a cell's value through. */
  styles(): Styles
  /** Whether any edit has been recorded, so `toBytes` can skip an untouched part. */
  changed(): boolean
  /** The amended styles.xml, byte-identical to the oracle's output. */
  serialize(): string
}

const prefixOf = (name: string): string => {
  const colon = name.indexOf(':')
  return colon === -1 ? '' : name.slice(0, colon + 1)
}

interface Table {
  readonly name: string
  readonly child: string
  present: boolean
  selfClosing: boolean
  openStart: number
  openEnd: number
  insertAt: number
  closeEnd: number
  /** Prefix appended child elements carry, and the container itself when created. */
  prefix: string
  /** Current elements: original, then any reserved seed, then appended. */
  readonly elements: string[]
  readonly map: Map<string, number>
  /** How many leading elements are original or reserved rather than appended. */
  baseCount: number
  seeded: boolean
  touched: boolean
  firstTouchSeq: number
}

interface RawContainer {
  present: boolean
  selfClosing: boolean
  openStart: number
  openEnd: number
  insertAt: number
  closeEnd: number
  prefix: string
  elements: string[]
}

/** Reads one container's direct children, keeping each whole, plus the offsets a
 * splice needs. Mirrors `readTable` in styles-writer, with the close-tag end
 * captured so a seeded container can be replaced in place. */
function readContainer(xml: string, container: string, child: string): RawContainer {
  let present = false
  let selfClosing = false
  let openStart = -1
  let openEnd = -1
  let insertAt = -1
  let closeEnd = -1
  let prefix = ''
  let inside = false
  let openElement = -1
  const elements: string[] = []

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === container) {
      present = true
      prefix = prefixOf(event.name)
      openStart = event.start
      openEnd = event.end
      selfClosing = event.selfClosing
      inside = !event.selfClosing
      if (selfClosing) {
        insertAt = event.end
        closeEnd = event.end
      }
      continue
    }
    if (event.kind === 'close' && event.localName === container) {
      insertAt = event.start
      closeEnd = event.end
      inside = false
      continue
    }
    if (inside && event.kind === 'open' && event.localName === child) {
      if (event.selfClosing) elements.push(xml.slice(event.start, event.end))
      else openElement = event.start
      continue
    }
    if (inside && event.kind === 'close' && event.localName === child && openElement !== -1) {
      elements.push(xml.slice(openElement, event.end))
      openElement = -1
    }
  }

  return { present, selfClosing, openStart, openEnd, insertAt, closeEnd, prefix, elements }
}

const malformed = (): XlsxError =>
  new XlsxError('malformed-xml', 'Style table has no styleSheet element', {
    part: 'xl/styles.xml',
  })

export function createStylesSession(original: string): StylesSession {
  const prefix = tablePrefixOf(original)

  // The styleSheet's own span, so a fresh numFmts (right after the open) or a
  // fresh fonts/fills/borders/cellXfs (before the close) lands where the oracle
  // creates it. Left at -1 when absent, and only a creation that needs it throws.
  let styleSheetOpenEnd = -1
  let styleSheetCloseStart = -1
  let insertPrefix = ''
  for (const event of readXml(original)) {
    if (event.kind === 'open' && event.localName === 'styleSheet' && styleSheetOpenEnd === -1) {
      styleSheetOpenEnd = event.end
    }
    if (event.kind === 'close' && event.localName === 'styleSheet') {
      styleSheetCloseStart = event.start
      insertPrefix = prefixOf(event.name)
    }
  }

  const makeTable = (name: string, child: string): Table => {
    const raw = readContainer(original, name, child)
    return {
      name,
      child,
      present: raw.present,
      selfClosing: raw.selfClosing,
      openStart: raw.openStart,
      openEnd: raw.openEnd,
      insertAt: raw.insertAt,
      closeEnd: raw.closeEnd,
      prefix: raw.present ? raw.prefix : insertPrefix,
      elements: raw.elements,
      map: firstOccurrenceMap(raw.elements),
      baseCount: raw.elements.length,
      seeded: false,
      touched: false,
      firstTouchSeq: 0,
    }
  }

  const fonts = makeTable('fonts', 'font')
  const fills = makeTable('fills', 'fill')
  const borders = makeTable('borders', 'border')
  const cellXfs = makeTable('cellXfs', 'xf')

  // numFmts is never looked up by element index, only through the id->code map, so
  // its original elements are left in the string and only the count is tracked.
  const numFmtsRaw = readContainer(original, 'numFmts', 'numFmt')
  const numFmts: Table = {
    name: 'numFmts',
    child: 'numFmt',
    present: numFmtsRaw.present,
    selfClosing: numFmtsRaw.selfClosing,
    openStart: numFmtsRaw.openStart,
    openEnd: numFmtsRaw.openEnd,
    insertAt: numFmtsRaw.insertAt,
    closeEnd: numFmtsRaw.closeEnd,
    prefix,
    elements: [],
    map: new Map(),
    baseCount: numFmtsRaw.elements.length,
    seeded: false,
    touched: false,
    firstTouchSeq: 0,
  }

  const parsed = readStyles(original)
  const numberFormats = new Map(parsed.numberFormats)
  const cellFormats = [...parsed.cellFormats]
  const stylesView: Styles = { numberFormats, cellFormats }
  const used = usedFormatIds(original)

  let touchCounter = 0
  let undo: Array<() => void> | undefined
  let formattingCache: Map<number, CellFormatting> | undefined

  const record = (action: () => void): void => {
    if (undo !== undefined) undo.push(action)
  }

  const markTouched = (table: Table): void => {
    if (table.touched) return
    table.touched = true
    table.firstTouchSeq = touchCounter++
    record(() => {
      table.touched = false
      touchCounter--
    })
  }

  const invalidateFormatting = (): void => {
    formattingCache = undefined
  }

  const appendElement = (table: Table, element: string, isXf: boolean): number => {
    const prefixed = element.replace(/^<[^\s/>]+/, `<${table.prefix}${table.child}`)
    const existing = table.map.get(prefixed)
    if (existing !== undefined) return existing
    markTouched(table)
    invalidateFormatting()
    const id = table.elements.length
    table.elements.push(prefixed)
    table.map.set(prefixed, id)
    let pushedFormat = false
    if (isXf) {
      cellFormats.push(attrId(prefixed, 'numFmtId'))
      pushedFormat = true
    }
    record(() => {
      table.elements.pop()
      if (table.map.get(prefixed) === id) table.map.delete(prefixed)
      if (pushedFormat) cellFormats.pop()
    })
    return id
  }

  const seedTable = (table: Table, reserved: readonly string[]): void => {
    markTouched(table)
    invalidateFormatting()
    table.seeded = true
    table.baseCount = reserved.length
    for (const element of reserved) {
      const id = table.elements.length
      table.elements.push(element)
      if (!table.map.has(element)) table.map.set(element, id)
    }
    record(() => {
      table.seeded = false
      table.baseCount = 0
      table.elements.length = 0
      table.map.clear()
    })
  }

  const seedFonts = (): void => {
    if (fonts.elements.length === 0) seedTable(fonts, [`<${fonts.prefix}font/>`])
  }
  const seedFills = (): void => {
    if (fills.elements.length === 0) {
      const p = fills.prefix
      const reserved = RESERVED_FILLS.replace(/<(\/?)(fill|patternFill)/g, `<$1${p}$2`).match(
        /<(?:[A-Za-z0-9]+:)?fill>[\s\S]*?<\/(?:[A-Za-z0-9]+:)?fill>/g,
      )
      seedTable(fills, reserved ?? [])
    }
  }
  const seedBorders = (): void => {
    if (borders.elements.length === 0) {
      const p = borders.prefix
      const reserved = RESERVED_BORDER.replace(
        /<(\/?)(border|left|right|top|bottom|diagonal)/g,
        `<$1${p}$2`,
      )
      seedTable(borders, [reserved])
    }
  }

  const idFor = (basedOn: number | undefined, attribute: string): number => {
    if (basedOn === undefined) return 0
    const element = cellXfs.elements[basedOn]
    return element === undefined ? 0 : attrId(element, attribute)
  }

  const applyCellFormat = (basedOn: number | undefined, overrides: CellOverrides): number => {
    const base = basedOn === undefined ? DEFAULT_XF : (cellXfs.elements[basedOn] ?? DEFAULT_XF)
    return appendElement(cellXfs, withOverrides(base, overrides), true)
  }

  const numberFormat = (
    basedOn: number | undefined,
    code: string,
    location: XlsxErrorContext,
  ): number => {
    const unwritable = findUnwritableCharacter(code)
    if (unwritable !== undefined) {
      throw new XlsxError(
        'unwritable-value',
        `Number format holds ${unwritable}, which cannot be written to xml`,
        { part: 'xl/styles.xml', ...location },
      )
    }

    if (basedOn !== undefined && numberFormatOf(stylesView, basedOn) === code) return basedOn
    if (basedOn === undefined) {
      for (let index = 0; index < cellFormats.length; index++) {
        if (numberFormatOf(stylesView, index) === code && isPlainFormat(cellXfs.elements[index])) {
          return index
        }
      }
    }

    let formatId = builtInFormatId(code)
    if (formatId === undefined) {
      for (const [id, existing] of numberFormats) if (existing === code) formatId = id
    }
    if (formatId === undefined) {
      formatId = FIRST_CUSTOM_FORMAT_ID
      while (used.has(formatId)) formatId++
      allocateNumberFormat(formatId, code)
    }
    return applyCellFormat(basedOn, { numFmtId: formatId })
  }

  const allocateNumberFormat = (id: number, code: string): void => {
    numberFormats.set(id, code)
    used.add(id)
    record(() => {
      numberFormats.delete(id)
      used.delete(id)
    })
    appendElement(
      numFmts,
      `<${prefix}numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`,
      false,
    )
  }

  const dateStyle = (basedOn: number | undefined): number => {
    if (basedOn !== undefined && isDateFormat(stylesView, basedOn)) return basedOn
    if (basedOn === undefined) {
      for (let index = 0; index < cellFormats.length; index++) {
        if (isDateFormat(stylesView, index) && isPlainFormat(cellXfs.elements[index])) return index
      }
    }
    return applyCellFormat(basedOn, { numFmtId: SHORT_DATE_FORMAT_ID })
  }

  const font = (
    basedOn: number | undefined,
    asked: FontFormat,
    location: XlsxErrorContext,
  ): number => {
    seedFonts()
    const current = parseFont(fonts.elements[idFor(basedOn, 'fontId')] ?? '')
    const merged: FontFormat = {
      bold: asked.bold ?? current.bold,
      italic: asked.italic ?? current.italic,
      strike: asked.strike ?? current.strike,
      underline: asked.underline ?? current.underline,
      verticalAlign: asked.verticalAlign ?? current.verticalAlign,
      size: asked.size ?? current.size,
      color: asked.color ?? current.color,
      name: asked.name ?? current.name,
    }
    const element = withNamespacePrefix(buildFontElement(merged, location), prefix)
    const id = appendElement(fonts, element, false)
    return applyCellFormat(basedOn, { fontId: id })
  }

  const fill = (
    basedOn: number | undefined,
    asked: FillFormat,
    location: XlsxErrorContext,
  ): number => {
    seedFills()
    const element = withNamespacePrefix(buildFillElement(asked, location), prefix)
    const id = appendElement(fills, element, false)
    return applyCellFormat(basedOn, { fillId: id })
  }

  const border = (
    basedOn: number | undefined,
    asked: BorderFormat,
    location: XlsxErrorContext,
  ): number => {
    seedBorders()
    const element = borders.elements[idFor(basedOn, 'borderId')]
    const current: Sides = element === undefined ? {} : parseBorder(element)
    const merged: {
      left?: Side
      right?: Side
      top?: Side
      bottom?: Side
      diagonal?: Side
      diagonalUp?: boolean
      diagonalDown?: boolean
    } = {}
    for (const name of SIDE_NAMES) {
      const side = asked[name] ?? asked.all ?? current[name]
      if (side !== undefined) merged[name] = side
    }
    if (asked.diagonal !== undefined) {
      const { up, down } = asked.diagonal
      merged.diagonal = { style: asked.diagonal.style, color: asked.diagonal.color }
      merged.diagonalUp = up === true
      merged.diagonalDown = down === true || (up === undefined && down === undefined)
    } else {
      merged.diagonal = current.diagonal
      merged.diagonalUp = current.diagonalUp
      merged.diagonalDown = current.diagonalDown
    }
    const builtElement = withNamespacePrefix(buildBorderElement(merged, location), prefix)
    const id = appendElement(borders, builtElement, false)
    return applyCellFormat(basedOn, { borderId: id })
  }

  const alignment = (
    basedOn: number | undefined,
    asked: Alignment,
    location: XlsxErrorContext,
  ): number => {
    validateAlignment(asked, location)
    const base = basedOn === undefined ? DEFAULT_XF : (cellXfs.elements[basedOn] ?? DEFAULT_XF)
    const current = parseAlignment(base)
    const merged: AlignmentAttributes = {
      horizontal: asked.horizontal ?? current.horizontal,
      vertical: asked.vertical ?? current.vertical,
      wrapText: asked.wrapText ?? current.wrapText,
      textRotation: asked.textRotation ?? current.textRotation,
      indent: asked.indent ?? current.indent,
    }
    const wanted = withAlignmentChild(base, buildAlignmentElement(merged, prefix), prefix)
    return appendElement(cellXfs, wanted, true)
  }

  const protection = (basedOn: number | undefined, asked: CellProtection): number => {
    const base = basedOn === undefined ? DEFAULT_XF : (cellXfs.elements[basedOn] ?? DEFAULT_XF)
    const current = parseProtection(base)
    const merged: ProtectionAttributes = {
      locked: asked.locked ?? current.locked,
      hidden: asked.hidden ?? current.hidden,
    }
    const wanted = withProtectionChild(base, buildProtectionElement(merged, prefix), prefix)
    return appendElement(cellXfs, wanted, true)
  }

  const transaction = <T>(body: () => T): T => {
    if (undo !== undefined) return body()
    undo = []
    try {
      const result = body()
      undo = undefined
      return result
    } catch (error) {
      const log = undo ?? []
      undo = undefined
      for (let index = log.length - 1; index >= 0; index--) log[index]?.()
      invalidateFormatting()
      throw error
    }
  }

  const formattingOf = (index: number): CellFormatting => {
    const xf = cellXfs.elements[index]
    if (xf === undefined) return {}
    formattingCache ??= new Map()
    const cached = formattingCache.get(index)
    if (cached !== undefined) return cached
    const resolved = formattingOfXf(xf, fonts.elements, fills.elements, borders.elements)
    formattingCache.set(index, resolved)
    return resolved
  }

  const changed = (): boolean =>
    fonts.touched || fills.touched || borders.touched || cellXfs.touched || numFmts.touched

  const presentSplice = (
    table: Table,
  ): { start: number; end: number; text: string; order: number } => {
    const count = table.elements.length
    if (table.seeded) {
      return {
        start: table.openStart,
        end: table.closeEnd,
        text: `<${table.prefix}${table.name} count="${count}">${table.elements.join('')}</${table.prefix}${table.name}>`,
        order: 0,
      }
    }
    if (table.selfClosing) {
      const opened = withAttribute(
        original.slice(table.openStart, table.openEnd - 2),
        'count',
        count,
      )
      return {
        start: table.openStart,
        end: table.openEnd,
        text: `${opened}>${table.elements.join('')}</${table.prefix}${table.name}>`,
        order: 0,
      }
    }
    const open = withAttribute(original.slice(table.openStart, table.openEnd), 'count', count)
    const interior = original.slice(table.openEnd, table.insertAt)
    const appended = table.elements.slice(table.baseCount).join('')
    return {
      start: table.openStart,
      end: table.insertAt,
      text: open + interior + appended,
      order: 0,
    }
  }

  const serialize = (): string => {
    // `order` breaks a tie between two zero-width inserts at the same offset — a
    // fresh numFmts (opened at styleSheet.end) and a fresh fonts/fills/borders/
    // cellXfs group (opened before the close) collide when the styleSheet is
    // empty. numFmts opens first in the document, so it must be applied last to
    // land ahead of the group.
    const splices: Array<{ start: number; end: number; text: string; order: number }> = []

    if (numFmts.touched) {
      const count = numFmts.baseCount + numFmts.elements.length
      const body = numFmts.elements.join('')
      if (numFmts.present) {
        if (numFmts.selfClosing) {
          const opened = withAttribute(
            original.slice(numFmts.openStart, numFmts.openEnd - 2),
            'count',
            count,
          )
          splices.push({
            start: numFmts.openStart,
            end: numFmts.openEnd,
            text: `${opened}>${body}</${prefix}numFmts>`,
            order: 0,
          })
        } else {
          const open = withAttribute(
            original.slice(numFmts.openStart, numFmts.openEnd),
            'count',
            count,
          )
          const interior = original.slice(numFmts.openEnd, numFmts.insertAt)
          splices.push({
            start: numFmts.openStart,
            end: numFmts.insertAt,
            text: open + interior + body,
            order: 0,
          })
        }
      } else {
        if (styleSheetOpenEnd === -1) throw malformed()
        splices.push({
          start: styleSheetOpenEnd,
          end: styleSheetOpenEnd,
          text: `<${prefix}numFmts count="${count}">${body}</${prefix}numFmts>`,
          // Applied last so it lands ahead of a fresh table group at the same spot.
          order: 1,
        })
      }
    }

    const created: Table[] = []
    for (const table of [fonts, fills, borders, cellXfs]) {
      if (!table.touched) continue
      if (table.present) splices.push(presentSplice(table))
      else created.push(table)
    }
    if (created.length > 0) {
      const nonXf = created
        .filter((table) => table !== cellXfs)
        .sort((a, b) => a.firstTouchSeq - b.firstTouchSeq)
      const ordered = created.includes(cellXfs) ? [...nonXf, cellXfs] : nonXf
      const position = cellXfs.present ? cellXfs.openStart : styleSheetCloseStart
      if (position === -1) throw malformed()
      const text = ordered
        .map(
          (table) =>
            `<${table.prefix}${table.name} count="${table.elements.length}">${table.elements.join('')}</${table.prefix}${table.name}>`,
        )
        .join('')
      splices.push({ start: position, end: position, text, order: 0 })
    }

    if (splices.length === 0) return original
    // Descending by start so a splice never shifts an earlier one's offsets. When
    // two share a start — a fresh numFmts inserted at styleSheet.end, right where
    // fonts opens, or a fresh table inserted where cellXfs opens — the wider
    // replace is applied first (end descending), then the lower-order insert last,
    // so each zero-width insert lands on the side the oracle put it on.
    splices.sort((a, b) => b.start - a.start || b.end - a.end || a.order - b.order)
    let result = original
    for (const splice of splices) {
      result = result.slice(0, splice.start) + splice.text + result.slice(splice.end)
    }
    return result
  }

  return {
    numberFormat,
    dateStyle,
    font,
    fill,
    border,
    alignment,
    protection,
    transaction,
    numberFormatOf: (index) => numberFormatOf(stylesView, index),
    formattingOf,
    styles: () => stylesView,
    changed,
    serialize,
  }
}

function firstOccurrenceMap(elements: readonly string[]): Map<string, number> {
  const map = new Map<string, number>()
  for (let index = 0; index < elements.length; index++) {
    const element = elements[index]
    if (element !== undefined && !map.has(element)) map.set(element, index)
  }
  return map
}
