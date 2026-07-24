import { XlsxError } from './errors.js'
import { builtInFormatId, isDateFormat, numberFormatOf, readStyles } from './styles.js'
import { findUnwritableCharacter, readXml, withAttribute } from './xml.js'

/** Built in short date format, so no custom numFmt has to be written. */
const SHORT_DATE_FORMAT_ID = 14

/** Custom format ids start here; everything below is built in. */
const FIRST_CUSTOM_FORMAT_ID = 164

const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export interface DateStyle {
  readonly xml: string
  /** Index a cell should carry in its `s` attribute to display as a date. */
  readonly index: number
}

const prefixOf = (name: string) => {
  const colon = name.indexOf(':')
  return colon === -1 ? '' : name.slice(0, colon + 1)
}

/** Prefixes every element name in a freshly built fragment. The builders emit
 *  plain names, and `ensureInTable` fixes only the outer open tag, so without
 *  this a fragment added to a prefixed table closed and nested unprefixed. */
const withNamespacePrefix = (fragment: string, prefix: string): string =>
  prefix === '' ? fragment : fragment.replace(/<(\/?)([A-Za-z])/g, `<$1${prefix}$2`)

/** One of the styles sub-tables — `fonts`, `fills`, `borders`, `cellXfs`. */
interface StyleTable {
  readonly elements: string[]
  /** Namespace prefix the file writes its elements with, `x:` or empty. */
  readonly prefix: string
  readonly openTag: string
  readonly openStart: number
  readonly openEnd: number
  readonly insertAt: number
  readonly selfClosing: boolean
}

/** Reads the direct `child` elements of `container`, each kept whole so a clone
 * carries its own children (an xf's alignment, a font's parts) with it. */
function readTable(xml: string, container: string, child: string): StyleTable | undefined {
  let openTag = ''
  let openStart = -1
  let openEnd = -1
  let insertAt = -1
  let selfClosing = false
  let inside = false
  let prefix = ''
  let openElement = -1
  const elements: string[] = []

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === container) {
      prefix = prefixOf(event.name)
      openTag = xml.slice(event.start, event.end)
      openStart = event.start
      openEnd = event.end
      selfClosing = event.selfClosing
      inside = !event.selfClosing
      if (selfClosing) insertAt = event.end
      continue
    }
    if (event.kind === 'close' && event.localName === container) {
      insertAt = event.start
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

  if (openStart === -1) return undefined
  return { elements, prefix, openTag, openStart, openEnd, insertAt, selfClosing }
}

/** Where a new sub-table is opened: before `cellXfs` so fonts/fills/borders keep
 * their place ahead of it, or before the closing tag when there is no cellXfs. */
function tableInsertPoint(xml: string): { position: number; prefix: string } {
  let cellXfsStart = -1
  let closeStart = -1
  let prefix = ''
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'cellXfs' && cellXfsStart === -1) {
      cellXfsStart = event.start
    }
    if (event.kind === 'close' && event.localName === 'styleSheet') {
      closeStart = event.start
      prefix = prefixOf(event.name)
    }
  }
  if (closeStart === -1) {
    throw new XlsxError('malformed-xml', 'Style table has no styleSheet element', {
      part: 'xl/styles.xml',
    })
  }
  return { position: cellXfsStart === -1 ? closeStart : cellXfsStart, prefix }
}

/** Adds `element` to `container`, or returns the index of an identical one. */
function ensureInTable(
  xml: string,
  container: string,
  child: string,
  element: string,
): { xml: string; id: number } {
  const table = readTable(xml, container, child)
  if (table === undefined) {
    const { position, prefix } = tableInsertPoint(xml)
    const prefixed = element.replace(/^<[^\s/>]+/, `<${prefix}${child}`)
    const created = `<${prefix}${container} count="1">${prefixed}</${prefix}${container}>`
    return { xml: xml.slice(0, position) + created + xml.slice(position), id: 0 }
  }

  const prefixed = element.replace(/^<[^\s/>]+/, `<${table.prefix}${child}`)
  const existing = table.elements.indexOf(prefixed)
  if (existing !== -1) return { xml, id: existing }

  const id = table.elements.length
  const openTag = withAttribute(table.openTag, 'count', id + 1)

  if (table.selfClosing) {
    const opened = `${openTag.slice(0, -2)}>`
    return {
      xml:
        xml.slice(0, table.openStart) +
        `${opened}${prefixed}</${table.prefix}${container}>` +
        xml.slice(table.openEnd),
      id,
    }
  }

  const head = xml.slice(0, table.openStart) + openTag
  const body = xml.slice(table.openEnd, table.insertAt)
  return { xml: `${head}${body}${prefixed}${xml.slice(table.insertAt)}`, id }
}

const DEFAULT_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

/** The parts of a cell format this library sets, each with its `apply*` flag. */
interface CellOverrides {
  readonly numFmtId?: number
  readonly fontId?: number
  readonly fillId?: number
  readonly borderId?: number
}

const OVERRIDE_FLAGS: ReadonlyArray<readonly [keyof CellOverrides, string]> = [
  ['numFmtId', 'applyNumberFormat'],
  ['fontId', 'applyFont'],
  ['fillId', 'applyFill'],
  ['borderId', 'applyBorder'],
]

/** Turns an `apply*` flag on, whether it was off, absent, or written as a word. */
function withApplyFlag(openTag: string, flag: string): string {
  const pattern = new RegExp(`${flag}=("|')(?:\\d+|true|false)\\1`)
  if (pattern.test(openTag)) return openTag.replace(pattern, `${flag}="1"`)
  return openTag.replace(/\/?>$/, (tag) => (tag === '/>' ? ` ${flag}="1"/>` : ` ${flag}="1">`))
}

function withOverrides(element: string, overrides: CellOverrides): string {
  const openTagEnd = element.indexOf('>') + 1
  let openTag = element.slice(0, openTagEnd)
  const rest = element.slice(openTagEnd)

  for (const [attribute, flag] of OVERRIDE_FLAGS) {
    const value = overrides[attribute]
    // The element may be written <xf> or <x:xf>, so the name is matched rather
    // than assumed. Missing it left the cell on General while it held a serial.
    if (value !== undefined) openTag = withApplyFlag(withAttribute(openTag, attribute, value), flag)
  }

  return `${openTag}${rest}`
}

/**
 * Finds a cell format that displays dates, adding one if the file has none.
 * `basedOn` keeps the font, fill and border of the cell being written to, so
 * applying a date format does not strip the rest of its formatting.
 */
export function ensureDateStyle(stylesXml: string, basedOn: number | undefined): DateStyle {
  const parsed = readStyles(stylesXml)
  if (basedOn !== undefined && isDateFormat(parsed, basedOn)) {
    return { xml: stylesXml, index: basedOn }
  }

  // A cell with no formatting of its own can borrow a date format already here,
  // but only from a plain xf: borrowing carries the whole xf, so a decorated one
  // would hand the blank cell a font, fill or border it never asked for.
  if (basedOn === undefined) {
    const xfs = readTable(stylesXml, 'cellXfs', 'xf')?.elements
    for (let index = 0; index < parsed.cellFormats.length; index++) {
      if (isDateFormat(parsed, index) && isPlainFormat(xfs?.[index])) {
        return { xml: stylesXml, index }
      }
    }
  }

  return applyCellFormat(stylesXml, basedOn, { numFmtId: SHORT_DATE_FORMAT_ID })
}

/** Puts a cell format carrying `overrides` in the table, reusing one if it fits. */
function applyCellFormat(
  stylesXml: string,
  basedOn: number | undefined,
  overrides: CellOverrides,
): DateStyle {
  const formats = readTable(stylesXml, 'cellXfs', 'xf')
  const base = basedOn === undefined ? DEFAULT_XF : (formats?.elements[basedOn] ?? DEFAULT_XF)
  const wanted = withOverrides(base, overrides)
  const { xml, id } = ensureInTable(stylesXml, 'cellXfs', 'xf', wanted)
  return { xml, index: id }
}

/** `true` and `'single'` both write a bare `<u/>`; the accounting and double
 * variants carry a `val`. A read reports `true` for a plain single underline. */
export type UnderlineStyle = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'

export type FontVerticalAlign = 'baseline' | 'superscript' | 'subscript'

/**
 * A colour a cell can carry: an `RRGGBB`/`AARRGGBB` hex literal, a reference into
 * the workbook theme's colour scheme (`tint` lightens or darkens it, -1 to 1), or
 * an index into the legacy palette. A theme or indexed colour is a reference, not
 * a value, so it is kept as one rather than flattened to the hex it resolves to.
 */
export type Color =
  | string
  | { readonly theme: number; readonly tint?: number }
  | { readonly indexed: number }

export interface FontFormat {
  readonly bold?: boolean
  readonly italic?: boolean
  readonly strike?: boolean
  readonly underline?: boolean | UnderlineStyle
  readonly verticalAlign?: FontVerticalAlign
  readonly size?: number
  readonly color?: Color
  readonly name?: string
}

// A plain <u/> is reported as the boolean true, so only the variants that carry a
// meaningful val are narrowed here; an unknown val falls back to a bare underline.
const REPORTED_UNDERLINES: ReadonlySet<UnderlineStyle> = new Set([
  'double',
  'singleAccounting',
  'doubleAccounting',
])

const toUnderline = (value: string | undefined): UnderlineStyle | undefined => {
  for (const known of REPORTED_UNDERLINES) if (known === value) return known
  return undefined
}

const VERTICAL_ALIGNS: ReadonlySet<FontVerticalAlign> = new Set([
  'baseline',
  'superscript',
  'subscript',
])

const toVertAlign = (value: string | undefined): FontVerticalAlign | undefined => {
  for (const known of VERTICAL_ALIGNS) if (known === value) return known
  return undefined
}

/** ECMA-376 stores a colour as eight hex digits, alpha first; a six-digit value
 * is the colour at full opacity. */
const HEX_COLOR = /^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/

export function normalizeColor(color: string): string {
  const hex = color.startsWith('#') ? color.slice(1) : color
  if (!HEX_COLOR.test(hex)) {
    throw new XlsxError('unwritable-value', `Colour "${color}" is not a 6 or 8 digit hex value`, {})
  }
  return (hex.length === 6 ? `FF${hex}` : hex).toUpperCase()
}

/** Reads a colour off a `<color>`, `<fgColor>` or `<bgColor>`, preferring an
 * explicit rgb, then a theme reference, then a palette index. A malformed rgb,
 * theme or index is dropped rather than reported, the way an unreadable
 * attribute is — a bad rgb read back would otherwise crash the next write that
 * re-serialises it. */
function parseColor(attributes: ReadonlyMap<string, string>): Color | undefined {
  const rgb = attributes.get('rgb')
  if (rgb !== undefined && HEX_COLOR.test(rgb)) return rgb
  const themeText = attributes.get('theme')
  if (themeText !== undefined) {
    const theme = Number(themeText)
    if (Number.isInteger(theme) && theme >= 0) {
      const tint = Number(attributes.get('tint'))
      return Number.isNaN(tint) ? { theme } : { theme, tint }
    }
  }
  const indexedText = attributes.get('indexed')
  if (indexedText !== undefined) {
    const indexed = Number(indexedText)
    if (Number.isInteger(indexed) && indexed >= 0) return { indexed }
  }
  return undefined
}

/** Serialises a colour to the attributes of a colour element, validating a
 * caller's value: a hex string, or a non-negative integer theme or index, with a
 * tint between -1 and 1. Returns a leading space so a site can inline it. */
function colorAttributes(color: Color): string {
  if (typeof color === 'string') return ` rgb="${normalizeColor(color)}"`
  if ('theme' in color) {
    if (!Number.isInteger(color.theme) || color.theme < 0) {
      throw new XlsxError(
        'unwritable-value',
        `Colour theme "${color.theme}" is not a non-negative integer`,
        {},
      )
    }
    if (color.tint === undefined) return ` theme="${color.theme}"`
    if (!Number.isFinite(color.tint) || color.tint < -1 || color.tint > 1) {
      throw new XlsxError(
        'unwritable-value',
        `Colour tint "${color.tint}" is not between -1 and 1`,
        {},
      )
    }
    return ` theme="${color.theme}" tint="${color.tint}"`
  }
  if (!Number.isInteger(color.indexed) || color.indexed < 0) {
    throw new XlsxError(
      'unwritable-value',
      `Colour index "${color.indexed}" is not a non-negative integer`,
      {},
    )
  }
  return ` indexed="${color.indexed}"`
}

/** A `<b val="0"/>` turns bold off, so presence alone is not the whole story. */
const flagOn = (value: string | undefined) => value !== '0' && value !== 'false'

function parseFont(element: string): FontFormat {
  const font: {
    bold?: boolean
    italic?: boolean
    strike?: boolean
    underline?: boolean | UnderlineStyle
    verticalAlign?: FontVerticalAlign
    size?: number
    color?: Color
    name?: string
  } = {}
  for (const event of readXml(element)) {
    if (event.kind !== 'open') continue
    switch (event.localName) {
      case 'b':
        font.bold = flagOn(event.attributes.get('val'))
        break
      case 'i':
        font.italic = flagOn(event.attributes.get('val'))
        break
      case 'strike':
        font.strike = flagOn(event.attributes.get('val'))
        break
      case 'u': {
        const val = event.attributes.get('val')
        // A cell can turn an inherited underline off; that is not underlined.
        if (val === 'none') break
        font.underline = toUnderline(val) ?? true
        break
      }
      case 'vertAlign': {
        const verticalAlign = toVertAlign(event.attributes.get('val'))
        if (verticalAlign !== undefined) font.verticalAlign = verticalAlign
        break
      }
      case 'sz': {
        const size = Number(event.attributes.get('val'))
        if (!Number.isNaN(size)) font.size = size
        break
      }
      case 'color': {
        const color = parseColor(event.attributes)
        if (color !== undefined) font.color = color
        break
      }
      case 'name': {
        const name = event.attributes.get('val')
        if (name !== undefined) font.name = name
        break
      }
    }
  }
  return font
}

function buildFontElement(font: FontFormat): string {
  let inner = ''
  if (font.bold === true) inner += '<b/>'
  if (font.italic === true) inner += '<i/>'
  if (font.strike === true) inner += '<strike/>'
  if (font.underline === true || font.underline === 'single') inner += '<u/>'
  else if (typeof font.underline === 'string') inner += `<u val="${font.underline}"/>`
  if (font.verticalAlign !== undefined) inner += `<vertAlign val="${font.verticalAlign}"/>`
  if (font.size !== undefined) {
    if (!Number.isFinite(font.size) || font.size <= 0) {
      throw new XlsxError('unwritable-value', `Font size ${font.size} is not a positive number`, {
        part: 'xl/styles.xml',
      })
    }
    inner += `<sz val="${font.size}"/>`
  }
  if (font.color !== undefined) inner += `<color${colorAttributes(font.color)}/>`
  if (font.name !== undefined) inner += `<name val="${escapeXml(font.name)}"/>`
  return `<font>${inner}</font>`
}

// readFormatting reads three of these off every xf, so the pattern per attribute
// is compiled once and reused rather than rebuilt on each read.
const numericAttributePatterns = new Map<string, RegExp>()
const numericAttributePattern = (attribute: string): RegExp => {
  const cached = numericAttributePatterns.get(attribute)
  if (cached !== undefined) return cached
  const pattern = new RegExp(`\\b${attribute}\\s*=\\s*["'](\\d+)["']`)
  numericAttributePatterns.set(attribute, pattern)
  return pattern
}

/** The value of a numeric xf attribute like `fontId`, or 0 (the default) when it names none. */
const attrId = (element: string, attribute: string): number => {
  const match = element.match(numericAttributePattern(attribute))
  return match?.[1] === undefined ? 0 : Number(match[1])
}

/** The id a cell format points at for one sub-table, or 0 (the default) when it names none. */
function idOf(stylesXml: string, basedOn: number | undefined, attribute: string): number {
  if (basedOn === undefined) return 0
  const element = readTable(stylesXml, 'cellXfs', 'xf')?.elements[basedOn]
  return element === undefined ? 0 : attrId(element, attribute)
}

/**
 * A cell format that decorates nothing: default font, fill and border and no
 * alignment or protection child. Only such a format is safe for a blank cell to
 * borrow, since a borrow carries the whole xf, not just its number format.
 */
function isPlainFormat(element: string | undefined): boolean {
  if (element === undefined) return true
  const idIsZero = (attribute: string) => attrId(element, attribute) === 0
  return (
    idIsZero('fontId') &&
    idIsZero('fillId') &&
    idIsZero('borderId') &&
    !/<[a-z0-9]*:?alignment[\s/>]/i.test(element) &&
    !/<[a-z0-9]*:?protection[\s/>]/i.test(element)
  )
}

/**
 * Seeds the reserved leading entries a plain cell points at (fontId 0, fillId 0
 * and 1, borderId 0) when the sub-table is absent — or present but empty, which
 * an odd file can be (`<fonts count="0"/>`). Without the seed a font, fill or
 * border we add lands on a reserved index and restyles every default cell.
 */
function withReservedTable(
  stylesXml: string,
  container: string,
  child: string,
  reservedInner: (prefix: string) => string,
  reservedCount: number,
): string {
  const table = readTable(stylesXml, container, child)
  if (table !== undefined && table.elements.length > 0) return stylesXml

  if (table === undefined) {
    const { position, prefix } = tableInsertPoint(stylesXml)
    const block = `<${prefix}${container} count="${reservedCount}">${reservedInner(prefix)}</${prefix}${container}>`
    return stylesXml.slice(0, position) + block + stylesXml.slice(position)
  }

  const { prefix } = table
  const block = `<${prefix}${container} count="${reservedCount}">${reservedInner(prefix)}</${prefix}${container}>`
  const end = table.selfClosing ? table.openEnd : table.insertAt + `</${prefix}${container}>`.length
  return stylesXml.slice(0, table.openStart) + block + stylesXml.slice(end)
}

// fontId 0 is the default font, so a file with no fonts (or an empty table) has
// one seeded rather than have the first font we add become every plain cell's.
function withReservedFont(stylesXml: string): string {
  return withReservedTable(stylesXml, 'fonts', 'font', (prefix) => `<${prefix}font/>`, 1)
}

/**
 * Applies `font` to a cell, merging onto the font it already has so setting bold
 * does not reset its size or colour. The merged font is added if the file has no
 * identical one, and a cell format pointing at it is returned.
 */
export function ensureFontStyle(
  stylesXml: string,
  basedOn: number | undefined,
  font: FontFormat,
): DateStyle {
  const seeded = withReservedFont(stylesXml)
  const fonts = readTable(seeded, 'fonts', 'font')
  const current = parseFont(fonts?.elements[idOf(seeded, basedOn, 'fontId')] ?? '')
  const merged: FontFormat = {
    bold: font.bold ?? current.bold,
    italic: font.italic ?? current.italic,
    strike: font.strike ?? current.strike,
    underline: font.underline ?? current.underline,
    verticalAlign: font.verticalAlign ?? current.verticalAlign,
    size: font.size ?? current.size,
    color: font.color ?? current.color,
    name: font.name ?? current.name,
  }

  const element = withNamespacePrefix(buildFontElement(merged), tablePrefix(seeded))
  const { xml, id } = ensureInTable(seeded, 'fonts', 'font', element)
  return applyCellFormat(xml, basedOn, { fontId: id })
}

/** ECMA-376 ST_PatternType, minus `none` (no fill) and `solid` (its own arm). */
export type PatternStyle =
  | 'gray125'
  | 'gray0625'
  | 'mediumGray'
  | 'darkGray'
  | 'lightGray'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis'

export interface SolidFill {
  readonly type: 'solid'
  /** The cell's background colour. */
  readonly color: Color
}

export interface PatternFill {
  readonly type: 'pattern'
  readonly pattern: PatternStyle
  /** The pattern's foreground colour. */
  readonly color: Color
  /** The colour behind the pattern; the default window background when absent. */
  readonly background?: Color
}

export type FillFormat = SolidFill | PatternFill

// fillId 0 (none) and 1 (gray125) are reserved, so a real solid fill is the
// third entry; a file with no fills table has these seeded before ours is added.
const RESERVED_FILLS =
  '<fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill>'

function withReservedFills(stylesXml: string): string {
  return withReservedTable(
    stylesXml,
    'fills',
    'fill',
    (prefix) => RESERVED_FILLS.replace(/<(\/?)(fill|patternFill)/g, `<$1${prefix}$2`),
    2,
  )
}

function buildFillElement(fill: FillFormat): string {
  const fg = `<fgColor${colorAttributes(fill.color)}/>`
  if (fill.type === 'solid') {
    return `<fill><patternFill patternType="solid">${fg}<bgColor indexed="64"/></patternFill></fill>`
  }
  const bg =
    fill.background === undefined
      ? '<bgColor indexed="64"/>'
      : `<bgColor${colorAttributes(fill.background)}/>`
  return `<fill><patternFill patternType="${fill.pattern}">${fg}${bg}</patternFill></fill>`
}

export function ensureFillStyle(
  stylesXml: string,
  basedOn: number | undefined,
  fill: FillFormat,
): DateStyle {
  const seeded = withReservedFills(stylesXml)
  const element = withNamespacePrefix(buildFillElement(fill), tablePrefix(seeded))
  const { xml, id } = ensureInTable(seeded, 'fills', 'fill', element)
  return applyCellFormat(xml, basedOn, { fillId: id })
}

/** ECMA-376 ST_BorderStyle; `none` is expressed by leaving the side out. */
export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot'

export interface BorderSide {
  readonly style: BorderStyle
  /** The side's line colour. */
  readonly color?: Color
}

export interface BorderFormat {
  readonly top?: BorderSide
  readonly bottom?: BorderSide
  readonly left?: BorderSide
  readonly right?: BorderSide
  /** Applied to all four sides; a specific side overrides it. */
  readonly all?: BorderSide
}

// The style read from a file is any ST_BorderStyle string, so the sides carried
// through a merge are looser than the BorderSide a caller passes in.
interface Side {
  readonly style: string
  readonly color?: Color
}
interface Sides {
  readonly left?: Side
  readonly right?: Side
  readonly top?: Side
  readonly bottom?: Side
}

const SIDE_NAMES = ['left', 'right', 'top', 'bottom'] as const

const RESERVED_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>'

function withReservedBorder(stylesXml: string): string {
  return withReservedTable(
    stylesXml,
    'borders',
    'border',
    (prefix) =>
      RESERVED_BORDER.replace(/<(\/?)(border|left|right|top|bottom|diagonal)/g, `<$1${prefix}$2`),
    1,
  )
}

function parseBorder(element: string): Sides {
  const sides: { left?: Side; right?: Side; top?: Side; bottom?: Side } = {}
  let inSide: 'left' | 'right' | 'top' | 'bottom' | undefined
  for (const event of readXml(element)) {
    if (event.kind === 'close') {
      if (inSide !== undefined && event.localName === inSide) inSide = undefined
      continue
    }
    if (event.kind !== 'open') continue
    if (event.localName === 'left' || event.localName === 'right') inSide = event.localName
    else if (event.localName === 'top' || event.localName === 'bottom') inSide = event.localName
    else {
      if (event.localName === 'color' && inSide !== undefined) {
        const color = parseColor(event.attributes)
        const style = sides[inSide]?.style
        if (color !== undefined && style !== undefined) sides[inSide] = { style, color }
      }
      continue
    }
    const style = event.attributes.get('style')
    if (style !== undefined) sides[inSide] = { style }
    if (event.selfClosing) inSide = undefined
  }
  return sides
}

const borderSidesAt = (stylesXml: string, id: number): Sides => {
  const element = readTable(stylesXml, 'borders', 'border')?.elements[id]
  return element === undefined ? {} : parseBorder(element)
}

const buildSide = (name: string, side: Side | undefined): string => {
  if (side === undefined) return `<${name}/>`
  if (side.color === undefined) return `<${name} style="${side.style}"/>`
  return `<${name} style="${side.style}"><color${colorAttributes(side.color)}/></${name}>`
}

const buildBorderElement = (sides: Sides): string =>
  `<border>${buildSide('left', sides.left)}${buildSide('right', sides.right)}${buildSide('top', sides.top)}${buildSide('bottom', sides.bottom)}<diagonal/></border>`

export function ensureBorderStyle(
  stylesXml: string,
  basedOn: number | undefined,
  border: BorderFormat,
): DateStyle {
  const seeded = withReservedBorder(stylesXml)
  const current = borderSidesAt(seeded, idOf(seeded, basedOn, 'borderId'))
  const merged: { left?: Side; right?: Side; top?: Side; bottom?: Side } = {}
  for (const name of SIDE_NAMES) {
    const side = border[name] ?? border.all ?? current[name]
    if (side !== undefined) merged[name] = side
  }

  const element = withNamespacePrefix(buildBorderElement(merged), tablePrefix(seeded))
  const { xml, id } = ensureInTable(seeded, 'borders', 'border', element)
  return applyCellFormat(xml, basedOn, { borderId: id })
}

export type HorizontalAlignment =
  | 'general'
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed'

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed'

export interface Alignment {
  readonly horizontal?: HorizontalAlignment
  readonly vertical?: VerticalAlignment
  readonly wrapText?: boolean
  /** Degrees anticlockwise, 0–180; 255 stacks the text top to bottom. */
  readonly textRotation?: number
  readonly indent?: number
}

// Unlike a font or border, alignment is not shared through a table: it lives as a
// child of the xf itself, read and written looser than the caller's Alignment
// because a file may hold a horizontal or vertical value outside our union.
interface AlignmentAttributes {
  readonly horizontal?: string
  readonly vertical?: string
  readonly wrapText?: boolean
  readonly textRotation?: number
  readonly indent?: number
}

const numberAttribute = (value: string | undefined): number | undefined => {
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isNaN(parsed) ? undefined : parsed
}

function parseAlignment(xf: string): AlignmentAttributes {
  const out: {
    horizontal?: string
    vertical?: string
    wrapText?: boolean
    textRotation?: number
    indent?: number
  } = {}
  for (const event of readXml(xf)) {
    if (event.kind !== 'open' || event.localName !== 'alignment') continue
    const horizontal = event.attributes.get('horizontal')
    if (horizontal !== undefined) out.horizontal = horizontal
    const vertical = event.attributes.get('vertical')
    if (vertical !== undefined) out.vertical = vertical
    const wrapText = event.attributes.get('wrapText')
    if (wrapText !== undefined) out.wrapText = flagOn(wrapText)
    const rotation = numberAttribute(event.attributes.get('textRotation'))
    if (rotation !== undefined) out.textRotation = rotation
    const indent = numberAttribute(event.attributes.get('indent'))
    if (indent !== undefined) out.indent = indent
  }
  return out
}

function buildAlignmentElement(alignment: AlignmentAttributes, prefix: string): string {
  let attributes = ''
  if (alignment.horizontal !== undefined) attributes += ` horizontal="${alignment.horizontal}"`
  if (alignment.vertical !== undefined) attributes += ` vertical="${alignment.vertical}"`
  if (alignment.wrapText === true) attributes += ' wrapText="1"'
  if (alignment.textRotation !== undefined)
    attributes += ` textRotation="${alignment.textRotation}"`
  if (alignment.indent !== undefined) attributes += ` indent="${alignment.indent}"`
  return `<${prefix}alignment${attributes}/>`
}

/** Drops any alignment already on `xf`, so the merged one written next to it is
 *  the only one, and matches a prefixed or unprefixed element name. */
const ALIGNMENT_CHILD =
  /<(?:[A-Za-z0-9]+:)?alignment\b(?:[^>]*\/>|[\s\S]*?<\/(?:[A-Za-z0-9]+:)?alignment>)/

/** Puts `alignment` in the xf as its first child, ahead of any protection or
 *  extLst, and turns applyAlignment on. A self-closing xf is reopened, and the
 *  close tag carries the file's prefix so a prefixed table stays well formed. */
function withAlignmentChild(xf: string, alignment: string, prefix: string): string {
  const close = xf.indexOf('>')
  const selfClosing = xf.charAt(close - 1) === '/'
  const openTag = withApplyFlag(
    `${xf.slice(0, selfClosing ? close - 1 : close)}>`,
    'applyAlignment',
  )
  const body = selfClosing ? '' : xf.slice(close + 1, xf.lastIndexOf('</'))
  return `${openTag}${alignment}${body.replace(ALIGNMENT_CHILD, '')}</${prefix}xf>`
}

function validateAlignment(alignment: Alignment): void {
  const { textRotation, indent } = alignment
  if (
    textRotation !== undefined &&
    (!Number.isInteger(textRotation) ||
      textRotation < 0 ||
      (textRotation > 180 && textRotation !== 255))
  ) {
    throw new XlsxError('unwritable-value', `Text rotation ${textRotation} is not 0–180 or 255`, {
      part: 'xl/styles.xml',
    })
  }
  if (indent !== undefined && (!Number.isInteger(indent) || indent < 0)) {
    throw new XlsxError('unwritable-value', `Indent ${indent} is not a whole number of steps`, {
      part: 'xl/styles.xml',
    })
  }
}

/**
 * Applies `alignment` to a cell, merging onto the alignment it already has so
 * setting wrap does not reset its horizontal choice. The merged xf is added if
 * the file has no identical one, and its index is returned.
 */
export function ensureAlignmentStyle(
  stylesXml: string,
  basedOn: number | undefined,
  alignment: Alignment,
): DateStyle {
  validateAlignment(alignment)
  const prefix = tablePrefix(stylesXml)
  const base =
    basedOn === undefined
      ? DEFAULT_XF
      : (readTable(stylesXml, 'cellXfs', 'xf')?.elements[basedOn] ?? DEFAULT_XF)
  const current = parseAlignment(base)
  const merged: AlignmentAttributes = {
    horizontal: alignment.horizontal ?? current.horizontal,
    vertical: alignment.vertical ?? current.vertical,
    wrapText: alignment.wrapText ?? current.wrapText,
    textRotation: alignment.textRotation ?? current.textRotation,
    indent: alignment.indent ?? current.indent,
  }

  const wanted = withAlignmentChild(base, buildAlignmentElement(merged, prefix), prefix)
  const { xml, id } = ensureInTable(stylesXml, 'cellXfs', 'xf', wanted)
  return { xml, index: id }
}

export interface CellProtection {
  /** Whether the cell resists editing once the sheet is protected. */
  readonly locked?: boolean
  /** Whether the cell's formula is hidden once the sheet is protected. */
  readonly hidden?: boolean
}

interface ProtectionAttributes {
  readonly locked?: boolean
  readonly hidden?: boolean
}

function parseProtection(xf: string): ProtectionAttributes {
  const out: { locked?: boolean; hidden?: boolean } = {}
  for (const event of readXml(xf)) {
    if (event.kind !== 'open' || event.localName !== 'protection') continue
    const locked = event.attributes.get('locked')
    if (locked !== undefined) out.locked = flagOn(locked)
    const hidden = event.attributes.get('hidden')
    if (hidden !== undefined) out.hidden = flagOn(hidden)
  }
  return out
}

function buildProtectionElement(protection: ProtectionAttributes, prefix: string): string {
  let attributes = ''
  if (protection.locked !== undefined) attributes += ` locked="${protection.locked ? '1' : '0'}"`
  if (protection.hidden !== undefined) attributes += ` hidden="${protection.hidden ? '1' : '0'}"`
  return `<${prefix}protection${attributes}/>`
}

const PROTECTION_CHILD = /<(?:[A-Za-z0-9]+:)?protection\b[^>]*\/>/

/** Puts `protection` in the xf after any alignment and before any extLst, the
 *  order CT_Xf wants, dropping a protection it had and turning applyProtection on. */
function withProtectionChild(xf: string, protection: string, prefix: string): string {
  const close = xf.indexOf('>')
  const selfClosing = xf.charAt(close - 1) === '/'
  const openTag = withApplyFlag(
    `${xf.slice(0, selfClosing ? close - 1 : close)}>`,
    'applyProtection',
  )
  const body = (selfClosing ? '' : xf.slice(close + 1, xf.lastIndexOf('</'))).replace(
    PROTECTION_CHILD,
    '',
  )
  const extLst = body.match(/<(?:[A-Za-z0-9]+:)?extLst\b/)
  const at = extLst?.index ?? body.length
  return `${openTag}${body.slice(0, at)}${protection}${body.slice(at)}</${prefix}xf>`
}

export function ensureProtectionStyle(
  stylesXml: string,
  basedOn: number | undefined,
  protection: CellProtection,
): DateStyle {
  const prefix = tablePrefix(stylesXml)
  const base =
    basedOn === undefined
      ? DEFAULT_XF
      : (readTable(stylesXml, 'cellXfs', 'xf')?.elements[basedOn] ?? DEFAULT_XF)
  const current = parseProtection(base)
  const merged: ProtectionAttributes = {
    locked: protection.locked ?? current.locked,
    hidden: protection.hidden ?? current.hidden,
  }

  const wanted = withProtectionChild(base, buildProtectionElement(merged, prefix), prefix)
  const { xml, id } = ensureInTable(stylesXml, 'cellXfs', 'xf', wanted)
  return { xml, index: id }
}

function tablePrefix(xml: string): string {
  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    if (event.localName !== 'numFmts' && event.localName !== 'styleSheet') continue
    const colon = event.name.indexOf(':')
    return colon === -1 ? '' : event.name.slice(0, colon + 1)
  }
  return ''
}

/** Every numFmtId the file mentions, including ones only used by dxfs. */
function usedFormatIds(xml: string): Set<number> {
  const used = new Set<number>()
  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    const declared = event.attributes.get('numFmtId')
    if (declared === undefined) continue
    const id = Number(declared)
    if (!Number.isNaN(id)) used.add(id)
  }
  return used
}

/** Adds a numFmt element, creating the numFmts table when the file has none. */
function withNumberFormat(xml: string, id: number, code: string, prefix: string): string {
  const element = `<${prefix}numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`
  let openStart = -1
  let openEnd = -1
  // Counted rather than read off the open tag, because a file that omits count
  // or gets it wrong would otherwise have that error carried forward.
  let children = 0

  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'numFmts' && event.selfClosing) {
      const opened = `${withAttribute(xml.slice(event.start, event.end - 2), 'count', 1)}>`
      return `${xml.slice(0, event.start)}${opened}${element}</${prefix}numFmts>${xml.slice(event.end)}`
    }
    if (event.kind === 'open' && event.localName === 'numFmts') {
      openStart = event.start
      openEnd = event.end
      continue
    }
    if (openStart !== -1 && event.kind === 'open' && event.localName === 'numFmt') {
      children++
      continue
    }
    // openStart guards a close with no open, which leaves the table to be
    // created below rather than splicing against a position that was never set.
    if (event.kind === 'close' && event.localName === 'numFmts' && openStart !== -1) {
      // A count that disagrees with the children makes Excel call the file
      // unreadable and offer to repair it, which rewrites the whole package.
      const head =
        xml.slice(0, openStart) +
        withAttribute(xml.slice(openStart, openEnd), 'count', children + 1) +
        xml.slice(openEnd, event.start)
      return head + element + xml.slice(event.start)
    }
  }

  // No table at all, so one is opened directly after the root element.
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'styleSheet') continue
    const table = `<${prefix}numFmts count="1">${element}</${prefix}numFmts>`
    return xml.slice(0, event.end) + table + xml.slice(event.end)
  }

  throw new XlsxError('malformed-xml', 'Style table has no styleSheet element', {
    part: 'xl/styles.xml',
  })
}

/**
 * Finds a cell format that shows `formatCode`, adding the format and the cell
 * format when the file has neither. `basedOn` keeps the font, fill and border
 * of the cell being written to.
 */
export function ensureNumberFormat(
  stylesXml: string,
  basedOn: number | undefined,
  formatCode: string,
): DateStyle {
  const unwritable = findUnwritableCharacter(formatCode)
  if (unwritable !== undefined) {
    throw new XlsxError(
      'unwritable-value',
      `Number format holds ${unwritable}, which cannot be written to xml`,
      { part: 'xl/styles.xml' },
    )
  }

  const parsed = readStyles(stylesXml)
  if (basedOn !== undefined && numberFormatOf(parsed, basedOn) === formatCode) {
    return { xml: stylesXml, index: basedOn }
  }
  // Only a cell with no formatting of its own can borrow another cell's format;
  // otherwise the borrowed xf brings its font, fill and border along and the
  // target's are silently dropped.
  if (basedOn === undefined) {
    const xfs = readTable(stylesXml, 'cellXfs', 'xf')?.elements
    for (let index = 0; index < parsed.cellFormats.length; index++) {
      if (numberFormatOf(parsed, index) === formatCode && isPlainFormat(xfs?.[index])) {
        return { xml: stylesXml, index }
      }
    }
  }

  let formatId = builtInFormatId(formatCode)
  let withFormat = stylesXml
  const prefix = tablePrefix(stylesXml)

  if (formatId === undefined) {
    for (const [id, code] of parsed.numberFormats) {
      if (code === formatCode) formatId = id
    }
  }
  if (formatId === undefined) {
    const used = usedFormatIds(stylesXml)
    formatId = FIRST_CUSTOM_FORMAT_ID
    while (used.has(formatId)) formatId++
    withFormat = withNumberFormat(stylesXml, formatId, formatCode, prefix)
  }

  return applyCellFormat(withFormat, basedOn, { numFmtId: formatId })
}

// --- Reading a cell's formatting back ---

/** The font, fill, border and alignment a cell format resolves to, each absent
 * when the cell uses the default and so carries none of its own to report. */
export interface CellFormatting {
  readonly font?: FontFormat
  readonly fill?: FillFormat
  readonly border?: BorderFormat
  readonly alignment?: Alignment
  readonly protection?: CellProtection
}

const PATTERN_STYLES: ReadonlySet<PatternStyle> = new Set([
  'gray125',
  'gray0625',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
])

const toPatternStyle = (value: string | undefined): PatternStyle | undefined => {
  for (const known of PATTERN_STYLES) if (known === value) return known
  return undefined
}

function parseFill(element: string): FillFormat | undefined {
  let patternType: string | undefined
  let foreground: Color | undefined
  let background: Color | undefined
  for (const event of readXml(element)) {
    if (event.kind !== 'open') continue
    if (event.localName === 'patternFill') patternType = event.attributes.get('patternType')
    if (event.localName === 'fgColor') foreground = parseColor(event.attributes) ?? foreground
    if (event.localName === 'bgColor') {
      // An indexed bgColor is the default window background (commonly indexed="64"),
      // not a background the cell chose, so only an rgb or theme one is reported.
      const color = parseColor(event.attributes)
      if (color !== undefined && !(typeof color === 'object' && 'indexed' in color))
        background = color
    }
  }
  if (foreground === undefined) return undefined
  if (patternType === 'solid') return { type: 'solid', color: foreground }
  const pattern = toPatternStyle(patternType)
  if (pattern === undefined) return undefined
  return background === undefined
    ? { type: 'pattern', pattern, color: foreground }
    : { type: 'pattern', pattern, color: foreground, background }
}

const BORDER_STYLES: ReadonlySet<BorderStyle> = new Set([
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
])

/** Narrows a file's border style to the known set, so no assertion is needed. */
function toBorderStyle(value: string): BorderStyle | undefined {
  for (const known of BORDER_STYLES) if (known === value) return known
  return undefined
}

/** The default font is 0, so a non-zero id is a font the cell was given. */
function fontFrom(xf: string, fonts: readonly string[]): FontFormat | undefined {
  const id = attrId(xf, 'fontId')
  const element = fonts[id]
  if (id === 0 || element === undefined) return undefined
  const font = parseFont(element)
  return Object.keys(font).length === 0 ? undefined : font
}

const fillFrom = (xf: string, fills: readonly string[]): FillFormat | undefined => {
  const element = fills[attrId(xf, 'fillId')]
  return element === undefined ? undefined : parseFill(element)
}

function borderFrom(xf: string, borders: readonly string[]): BorderFormat | undefined {
  const element = borders[attrId(xf, 'borderId')]
  if (element === undefined) return undefined
  const sides = parseBorder(element)
  const border: { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide } =
    {}
  for (const name of SIDE_NAMES) {
    const side = sides[name]
    const style = side === undefined ? undefined : toBorderStyle(side.style)
    if (side === undefined || style === undefined) continue
    border[name] = side.color === undefined ? { style } : { style, color: side.color }
  }
  return Object.keys(border).length === 0 ? undefined : border
}

const HORIZONTAL_ALIGNMENTS: ReadonlySet<HorizontalAlignment> = new Set([
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
])

const VERTICAL_ALIGNMENTS: ReadonlySet<VerticalAlignment> = new Set([
  'top',
  'center',
  'bottom',
  'justify',
  'distributed',
])

const toHorizontal = (value: string | undefined): HorizontalAlignment | undefined => {
  for (const known of HORIZONTAL_ALIGNMENTS) if (known === value) return known
  return undefined
}

const toVertical = (value: string | undefined): VerticalAlignment | undefined => {
  for (const known of VERTICAL_ALIGNMENTS) if (known === value) return known
  return undefined
}

/** Alignment lives on the xf, so it is read straight off it, narrowed to the
 * known horizontal and vertical values the way a border style is. */
function alignmentFrom(xf: string): Alignment | undefined {
  const parsed = parseAlignment(xf)
  const horizontal = toHorizontal(parsed.horizontal)
  const vertical = toVertical(parsed.vertical)
  const alignment: {
    horizontal?: HorizontalAlignment
    vertical?: VerticalAlignment
    wrapText?: boolean
    textRotation?: number
    indent?: number
  } = {}
  if (horizontal !== undefined) alignment.horizontal = horizontal
  if (vertical !== undefined) alignment.vertical = vertical
  if (parsed.wrapText !== undefined) alignment.wrapText = parsed.wrapText
  if (parsed.textRotation !== undefined) alignment.textRotation = parsed.textRotation
  if (parsed.indent !== undefined) alignment.indent = parsed.indent
  return Object.keys(alignment).length === 0 ? undefined : alignment
}

/** Protection, like alignment, lives on the xf, so it is read straight off it. */
function protectionFrom(xf: string): CellProtection | undefined {
  const parsed = parseProtection(xf)
  return Object.keys(parsed).length === 0 ? undefined : parsed
}

/** Resolves every cell format's font, fill, border, alignment and protection in
 * one pass, so a read looks each up by the cell's `s` index. */
export function readFormatting(stylesXml: string): readonly CellFormatting[] {
  const xfs = readTable(stylesXml, 'cellXfs', 'xf')?.elements ?? []
  const fonts = readTable(stylesXml, 'fonts', 'font')?.elements ?? []
  const fills = readTable(stylesXml, 'fills', 'fill')?.elements ?? []
  const borders = readTable(stylesXml, 'borders', 'border')?.elements ?? []
  return xfs.map((xf) => {
    const font = fontFrom(xf, fonts)
    const fill = fillFrom(xf, fills)
    const border = borderFrom(xf, borders)
    const alignment = alignmentFrom(xf)
    const protection = protectionFrom(xf)
    return {
      ...(font === undefined ? {} : { font }),
      ...(fill === undefined ? {} : { fill }),
      ...(border === undefined ? {} : { border }),
      ...(alignment === undefined ? {} : { alignment }),
      ...(protection === undefined ? {} : { protection }),
    }
  })
}

const WRITABLE_UNDERLINES: ReadonlySet<UnderlineStyle> = new Set([
  'single',
  'double',
  'singleAccounting',
  'doubleAccounting',
])

const inUnion = <T>(known: ReadonlySet<T>, value: unknown): boolean => {
  for (const member of known) if (member === value) return true
  return false
}

/**
 * Refuses a style option outside its union at the `set()`/`format()` call. Types
 * stop a TypeScript caller; a JS caller, a JSON payload or an `any` at a boundary
 * reaches here, and an out-of-union enum was otherwise interpolated raw into
 * styles.xml — a bad value that broke the whole workbook, not just the one edit.
 */
export function checkStyleOptions(options: unknown, reference: string): void {
  if (typeof options !== 'object' || options === null) return

  const refuse = (what: string, value: unknown): never => {
    throw new XlsxError(
      'unwritable-value',
      `Cell ${reference} was given ${what} "${String(value)}", which is not one this library writes`,
      { part: 'xl/styles.xml', reference },
    )
  }

  if (
    'alignment' in options &&
    typeof options.alignment === 'object' &&
    options.alignment !== null
  ) {
    const alignment = options.alignment
    if ('horizontal' in alignment && alignment.horizontal !== undefined) {
      if (!inUnion(HORIZONTAL_ALIGNMENTS, alignment.horizontal))
        refuse('a horizontal alignment', alignment.horizontal)
    }
    if ('vertical' in alignment && alignment.vertical !== undefined) {
      if (!inUnion(VERTICAL_ALIGNMENTS, alignment.vertical))
        refuse('a vertical alignment', alignment.vertical)
    }
  }

  if ('font' in options && typeof options.font === 'object' && options.font !== null) {
    const font = options.font
    if ('underline' in font && typeof font.underline === 'string') {
      if (!inUnion(WRITABLE_UNDERLINES, font.underline))
        refuse('an underline style', font.underline)
    }
    if ('verticalAlign' in font && font.verticalAlign !== undefined) {
      if (!inUnion(VERTICAL_ALIGNS, font.verticalAlign))
        refuse('a font vertical alignment', font.verticalAlign)
    }
  }

  if ('border' in options && typeof options.border === 'object' && options.border !== null) {
    const border = options.border
    const side = (value: unknown, name: string): void => {
      if (typeof value === 'object' && value !== null && 'style' in value) {
        if (!inUnion(BORDER_STYLES, value.style)) refuse(`a ${name} border style`, value.style)
      }
    }
    if ('all' in border) side(border.all, 'all')
    if ('left' in border) side(border.left, 'left')
    if ('right' in border) side(border.right, 'right')
    if ('top' in border) side(border.top, 'top')
    if ('bottom' in border) side(border.bottom, 'bottom')
  }

  if ('fill' in options && typeof options.fill === 'object' && options.fill !== null) {
    const fill = options.fill
    const color = 'color' in fill ? fill.color : undefined
    if (color === undefined) refuse('a fill', 'with no colour')
    if ('type' in fill && fill.type === 'pattern' && 'pattern' in fill) {
      if (!inUnion(PATTERN_STYLES, fill.pattern)) refuse('a fill pattern', fill.pattern)
    }
  }
}
