import { indexToColumn } from './reference.js'
import { readXml, readXmlBytes } from './xml.js'

/** The page setup for printing: which way the paper turns and the print scale. */
export interface PageSetup {
  readonly orientation?: 'portrait' | 'landscape'
  /** Print scale as a whole percentage, 10 to 400. */
  readonly scale?: number
}

/** Whether a sheet prints its gridlines and its row and column headings. */
export interface PrintOptions {
  readonly gridlines?: boolean
  readonly headings?: boolean
}

/** The print margins, in inches. Every field is optional. */
export interface PageMargins {
  readonly left?: number
  readonly right?: number
  readonly top?: number
  readonly bottom?: number
  readonly header?: number
  readonly footer?: number
}

/**
 * One position of a header or footer. Each string is Excel's field-code text for
 * that position: `&P` the page number, `&N` the page count, `&D` the date, `&A`
 * the sheet name, `&&` a literal ampersand, `&"font,style"` a font. Plain text
 * needs no code.
 */
export interface HeaderFooterSection {
  readonly left?: string
  readonly center?: string
  readonly right?: string
}

/** A sheet's printed header and footer, each split into left, centre and right. */
export interface HeaderFooter {
  readonly header?: HeaderFooterSection
  readonly footer?: HeaderFooterSection
}

const MARGIN_KEYS = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const
// Excel's default margins, so a partial edit still writes a complete element.
const DEFAULT_MARGINS: Required<PageMargins> = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
}

// The worksheet children that follow pageMargins, in schema order, so a fresh
// element lands before the first one present and after everything before it.
const AFTER_PAGE_MARGINS =
  'pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF drawingHF picture oleObjects controls webPublishItems tableParts extLst'.split(
    ' ',
  )

const numberAttribute = (
  attributes: ReadonlyMap<string, string>,
  name: string,
): number | undefined => {
  const raw = attributes.get(name)
  if (raw === undefined) return undefined
  const value = Number(raw)
  return Number.isFinite(value) ? value : undefined
}

/** Sets or replaces an attribute on an open tag, keeping the rest. */
function setAttribute(openTag: string, name: string, value: string | number): string {
  const pattern = new RegExp(`(\\s${name}\\s*=\\s*)("[^"]*"|'[^']*')`)
  if (pattern.test(openTag)) return openTag.replace(pattern, `$1"${value}"`)
  return openTag.replace(/^<([^\s/>]+)/, `<$1 ${name}="${value}"`)
}

/** Inserts a fresh self-closing element before the first of `successors` present,
 * or before the worksheet's close when it has none. */
function insertBeforeSuccessor(
  sheetXml: string,
  element: string,
  successors: readonly string[],
): string {
  let at = sheetXml.indexOf('</worksheet')
  if (at === -1) at = sheetXml.length
  for (const successor of successors) {
    const found = sheetXml.search(new RegExp(`<(?:\\w+:)?${successor}\\b`))
    if (found !== -1 && found < at) at = found
  }
  return `${sheetXml.slice(0, at)}${element}${sheetXml.slice(at)}`
}

function selfClosingTag(
  sheetXml: string,
  localName: string,
): { start: number; end: number } | undefined {
  const match = new RegExp(`<(?:\\w+:)?${localName}\\b[^>]*/>`).exec(sheetXml)
  if (match === null) return undefined
  return { start: match.index, end: match.index + match[0].length }
}

/**
 * Writes the page margins, merging onto the ones the sheet has (or Excel's
 * defaults), so a partial edit still leaves a complete element. All six values
 * are written, in inches.
 */
export function withPageMargins(sheetXml: string, margins: PageMargins): string {
  const tag = selfClosingTag(sheetXml, 'pageMargins')
  const current: { -readonly [K in keyof Required<PageMargins>]: number } = { ...DEFAULT_MARGINS }
  if (tag !== undefined) {
    for (const event of readXml(sheetXml)) {
      if (event.kind === 'open' && event.localName === 'pageMargins') {
        for (const key of MARGIN_KEYS) {
          const value = numberAttribute(event.attributes, key)
          if (value !== undefined) current[key] = value
        }
      }
    }
  }
  for (const key of MARGIN_KEYS) {
    const value = margins[key]
    if (value !== undefined) current[key] = value
  }
  let attributes = ''
  for (const key of MARGIN_KEYS) attributes += ` ${key}="${current[key]}"`
  const element = `<pageMargins${attributes}/>`
  if (tag !== undefined) return sheetXml.slice(0, tag.start) + element + sheetXml.slice(tag.end)
  return insertBeforeSuccessor(sheetXml, element, AFTER_PAGE_MARGINS)
}

/**
 * Writes the page setup, setting only the orientation and scale a caller gives and
 * keeping any other attribute the sheet's `pageSetup` already carries.
 */
export function withPageSetup(sheetXml: string, setup: PageSetup): string {
  const tag = selfClosingTag(sheetXml, 'pageSetup')
  if (tag !== undefined) {
    let openTag = sheetXml.slice(tag.start, tag.end)
    if (setup.orientation !== undefined)
      openTag = setAttribute(openTag, 'orientation', setup.orientation)
    if (setup.scale !== undefined) openTag = setAttribute(openTag, 'scale', setup.scale)
    return sheetXml.slice(0, tag.start) + openTag + sheetXml.slice(tag.end)
  }
  const orientation = setup.orientation === undefined ? '' : ` orientation="${setup.orientation}"`
  const scale = setup.scale === undefined ? '' : ` scale="${setup.scale}"`
  return insertBeforeSuccessor(
    sheetXml,
    `<pageSetup${orientation}${scale}/>`,
    AFTER_PAGE_MARGINS.slice(1),
  )
}

// printOptions sits just before pageMargins in the worksheet schema.
const AFTER_PRINT_OPTIONS = ['pageMargins', ...AFTER_PAGE_MARGINS]

const flagAttribute = (
  attributes: ReadonlyMap<string, string>,
  name: string,
): boolean | undefined => {
  const raw = attributes.get(name)
  if (raw === undefined) return undefined
  return raw === '1' || raw === 'true'
}

/**
 * Writes the print options a caller gives, keeping any other attribute the
 * `printOptions` element carries. Turning gridlines on also sets `gridLinesSet`,
 * the flag Excel pairs with it.
 */
export function withPrintOptions(sheetXml: string, options: PrintOptions): string {
  const tag = selfClosingTag(sheetXml, 'printOptions')
  if (tag !== undefined) {
    let openTag = sheetXml.slice(tag.start, tag.end)
    if (options.gridlines !== undefined) {
      openTag = setAttribute(openTag, 'gridLines', options.gridlines ? 1 : 0)
      openTag = setAttribute(openTag, 'gridLinesSet', options.gridlines ? 1 : 0)
    }
    if (options.headings !== undefined) {
      openTag = setAttribute(openTag, 'headings', options.headings ? 1 : 0)
    }
    return sheetXml.slice(0, tag.start) + openTag + sheetXml.slice(tag.end)
  }
  let attributes = ''
  if (options.gridlines !== undefined) {
    const flag = options.gridlines ? 1 : 0
    attributes += ` gridLines="${flag}" gridLinesSet="${flag}"`
  }
  if (options.headings !== undefined) attributes += ` headings="${options.headings ? 1 : 0}"`
  return insertBeforeSuccessor(sheetXml, `<printOptions${attributes}/>`, AFTER_PRINT_OPTIONS)
}

/** Reads whether the sheet prints gridlines and headings, each off by default. */
export function readPrintOptions(bytes: Uint8Array): { gridlines: boolean; headings: boolean } {
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'printOptions') continue
    return {
      gridlines: flagAttribute(event.attributes, 'gridLines') ?? false,
      headings: flagAttribute(event.attributes, 'headings') ?? false,
    }
  }
  return { gridlines: false, headings: false }
}

/** Reads the page setup — the orientation and scale this models. */
export function readPageSetup(bytes: Uint8Array): PageSetup {
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'pageSetup') continue
    const orientation = event.attributes.get('orientation')
    const scale = numberAttribute(event.attributes, 'scale')
    const result: { -readonly [K in keyof PageSetup]?: PageSetup[K] } = {}
    if (orientation === 'portrait' || orientation === 'landscape') result.orientation = orientation
    if (scale !== undefined) result.scale = scale
    return result
  }
  return {}
}

/** Reads the page margins the sheet stores. */
export function readPageMargins(bytes: Uint8Array): PageMargins {
  for (const event of readXmlBytes(bytes)) {
    if (event.kind !== 'open' || event.localName !== 'pageMargins') continue
    const margins: { -readonly [K in keyof PageMargins]?: number } = {}
    for (const key of MARGIN_KEYS) {
      const value = numberAttribute(event.attributes, key)
      if (value !== undefined) margins[key] = value
    }
    return margins
  }
  return {}
}

const escapeXml = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

// The headerFooter children that follow oddHeader and oddFooter, so an inserted
// odd section lands before an even or first one the sheet already carries.
const HEADER_FOOTER_SUCCESSORS = AFTER_PAGE_MARGINS.slice(
  AFTER_PAGE_MARGINS.indexOf('headerFooter') + 1,
)

// A header or footer is one string, its three positions marked by &L, &C and &R.
// A doubled && is a literal ampersand, so it is not a position marker.
function buildSectionString(section: HeaderFooterSection): string {
  let out = ''
  if (section.left !== undefined) out += `&L${section.left}`
  if (section.center !== undefined) out += `&C${section.center}`
  if (section.right !== undefined) out += `&R${section.right}`
  return out
}

function parseSectionString(text: string): HeaderFooterSection {
  const result: { -readonly [K in keyof HeaderFooterSection]?: string } = {}
  let current: 'left' | 'center' | 'right' = 'center'
  let buffer = ''
  const flush = () => {
    if (buffer !== '') result[current] = buffer
    buffer = ''
  }
  for (let index = 0; index < text.length; index++) {
    const char = text[index]
    const next = text[index + 1]
    if (char === '&' && next === '&') {
      buffer += '&&'
      index++
    } else if (char === '&' && (next === 'L' || next === 'C' || next === 'R')) {
      flush()
      current = next === 'L' ? 'left' : next === 'C' ? 'center' : 'right'
      index++
    } else if (char === '&' && next !== undefined) {
      buffer += char + next
      index++
    } else {
      buffer += char
    }
  }
  flush()
  return result
}

function headerFooterElement(
  sheetXml: string,
): { start: number; end: number; open: string; inner: string; close: string } | undefined {
  const selfClosing = /<(?:\w+:)?headerFooter\b[^>]*\/>/.exec(sheetXml)
  if (selfClosing !== null) {
    return {
      start: selfClosing.index,
      end: selfClosing.index + selfClosing[0].length,
      open: selfClosing[0].replace(/\/>$/, '>'),
      inner: '',
      close: '</headerFooter>',
    }
  }
  const paired = /<(?:\w+:)?headerFooter\b[^>]*>([\s\S]*?)<\/(?:\w+:)?headerFooter>/.exec(sheetXml)
  if (paired === null) return undefined
  const openEnd = paired[0].indexOf('>') + 1
  const closeStart = paired[0].lastIndexOf('</')
  return {
    start: paired.index,
    end: paired.index + paired[0].length,
    open: paired[0].slice(0, openEnd),
    inner: paired[0].slice(openEnd, closeStart),
    close: paired[0].slice(closeStart),
  }
}

// Replaces the named child in a headerFooter's inner content, or inserts it in
// schema order: oddHeader first, oddFooter right after it.
function setHeaderFooterChild(inner: string, name: string, content: string): string {
  const element = `<${name}>${content}</${name}>`
  const existing = new RegExp(
    `<(?:\\w+:)?${name}\\b[^>]*>[\\s\\S]*?</(?:\\w+:)?${name}>|<(?:\\w+:)?${name}\\b[^>]*/>`,
  ).exec(inner)
  if (existing !== null) {
    return (
      inner.slice(0, existing.index) + element + inner.slice(existing.index + existing[0].length)
    )
  }
  if (name === 'oddFooter') {
    const header = /<(?:\w+:)?oddHeader\b[^>]*>[\s\S]*?<\/(?:\w+:)?oddHeader>/.exec(inner)
    if (header !== null) {
      const at = header.index + header[0].length
      return inner.slice(0, at) + element + inner.slice(at)
    }
  }
  return element + inner
}

/**
 * Writes the odd header and footer a caller gives, replacing that section whole
 * and leaving the other section, plus any even or first variants the sheet
 * carries, untouched.
 */
export function withHeaderFooter(sheetXml: string, headerFooter: HeaderFooter): string {
  const children: { name: string; content: string }[] = []
  if (headerFooter.header !== undefined)
    children.push({ name: 'oddHeader', content: buildSectionString(headerFooter.header) })
  if (headerFooter.footer !== undefined)
    children.push({ name: 'oddFooter', content: buildSectionString(headerFooter.footer) })

  const element = headerFooterElement(sheetXml)
  if (element === undefined) {
    let inner = ''
    for (const child of children)
      inner += `<${child.name}>${escapeXml(child.content)}</${child.name}>`
    return insertBeforeSuccessor(
      sheetXml,
      `<headerFooter>${inner}</headerFooter>`,
      HEADER_FOOTER_SUCCESSORS,
    )
  }
  let inner = element.inner
  for (const child of children)
    inner = setHeaderFooterChild(inner, child.name, escapeXml(child.content))
  return (
    sheetXml.slice(0, element.start) +
    element.open +
    inner +
    element.close +
    sheetXml.slice(element.end)
  )
}

/** Reads the odd header and footer, each parsed into its left, centre and right. */
export function readHeaderFooter(bytes: Uint8Array): HeaderFooter {
  const result: { -readonly [K in keyof HeaderFooter]?: HeaderFooterSection } = {}
  let capture: 'oddHeader' | 'oddFooter' | undefined
  let text = ''
  for (const event of readXmlBytes(bytes)) {
    if (
      event.kind === 'open' &&
      (event.localName === 'oddHeader' || event.localName === 'oddFooter')
    ) {
      if (event.selfClosing) continue
      capture = event.localName
      text = ''
    } else if (event.kind === 'text' && capture !== undefined) {
      text += event.text
    } else if (event.kind === 'close' && capture !== undefined && event.localName === capture) {
      const section = parseSectionString(text)
      if (Object.keys(section).length > 0) {
        if (capture === 'oddHeader') result.header = section
        else result.footer = section
      }
      capture = undefined
    }
  }
  return result
}

/** A sheet's manual page breaks: the rows and columns that begin a new page. */
export interface PageBreaks {
  /** One-based rows that begin a new page — the break sits above each. */
  readonly rows: readonly number[]
  /** Column letters that begin a new page — the break sits to the left of each. */
  readonly columns: readonly string[]
}

// A row break spans every column, a column break every row; the max is the last
// index, zero-based, of that span.
const LAST_COLUMN_INDEX = 16383
const LAST_ROW_INDEX = 1048575

const ROW_BREAKS_SUCCESSORS = AFTER_PAGE_MARGINS.slice(AFTER_PAGE_MARGINS.indexOf('rowBreaks') + 1)
const COL_BREAKS_SUCCESSORS = AFTER_PAGE_MARGINS.slice(AFTER_PAGE_MARGINS.indexOf('colBreaks') + 1)

function pairedElement(
  sheetXml: string,
  localName: string,
): { start: number; end: number; inner: string } | undefined {
  const selfClosing = new RegExp(`<(?:\\w+:)?${localName}\\b[^>]*/>`).exec(sheetXml)
  if (selfClosing !== null) {
    return { start: selfClosing.index, end: selfClosing.index + selfClosing[0].length, inner: '' }
  }
  const paired = new RegExp(
    `<(?:\\w+:)?${localName}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${localName}>`,
  ).exec(sheetXml)
  if (paired === null) return undefined
  const openEnd = paired[0].indexOf('>') + 1
  const closeStart = paired[0].lastIndexOf('</')
  return {
    start: paired.index,
    end: paired.index + paired[0].length,
    inner: paired[0].slice(openEnd, closeStart),
  }
}

// Existing breaks are kept verbatim, keyed by id, so their own attributes and any
// automatic ones survive. A break already at an id we are adding wins, so the same
// break added twice stays one.
function readBreaks(inner: string): Map<number, string> {
  const breaks = new Map<number, string>()
  for (const match of inner.matchAll(/<(?:\w+:)?brk\b[^>]*\/>/g)) {
    const idMatch = /\bid\s*=\s*"(\d+)"|\bid\s*=\s*'(\d+)'/.exec(match[0])
    const id = idMatch === null ? 0 : Number(idMatch[1] ?? idMatch[2])
    breaks.set(id, match[0])
  }
  return breaks
}

function withBreaks(
  sheetXml: string,
  localName: string,
  childMax: number,
  ids: readonly number[],
  successors: readonly string[],
): string {
  const element = pairedElement(sheetXml, localName)
  const breaks = element === undefined ? new Map<number, string>() : readBreaks(element.inner)
  for (const id of ids) {
    if (!breaks.has(id)) breaks.set(id, `<brk id="${id}" max="${childMax}" man="1"/>`)
  }
  let inner = ''
  let manual = 0
  for (const [, raw] of [...breaks.entries()].sort((a, b) => a[0] - b[0])) {
    inner += raw
    if (/\bman\s*=\s*["']1["']/.test(raw)) manual++
  }
  const container = `<${localName} count="${breaks.size}" manualBreakCount="${manual}">${inner}</${localName}>`
  if (element !== undefined) {
    return sheetXml.slice(0, element.start) + container + sheetXml.slice(element.end)
  }
  return insertBeforeSuccessor(sheetXml, container, successors)
}

/** Adds manual row page breaks, each `id` the zero-based row that begins a page. */
export function withRowBreaks(sheetXml: string, ids: readonly number[]): string {
  return withBreaks(sheetXml, 'rowBreaks', LAST_COLUMN_INDEX, ids, ROW_BREAKS_SUCCESSORS)
}

/** Adds manual column page breaks, each `id` the zero-based column that begins a page. */
export function withColumnBreaks(sheetXml: string, ids: readonly number[]): string {
  return withBreaks(sheetXml, 'colBreaks', LAST_ROW_INDEX, ids, COL_BREAKS_SUCCESSORS)
}

/** Reads the manual page breaks, rows as one-based numbers and columns as letters. */
export function readPageBreaks(bytes: Uint8Array): PageBreaks {
  const rowIds: number[] = []
  const columnIds: number[] = []
  let inside: 'rowBreaks' | 'colBreaks' | undefined
  for (const event of readXmlBytes(bytes)) {
    if (
      event.kind === 'open' &&
      (event.localName === 'rowBreaks' || event.localName === 'colBreaks')
    ) {
      inside = event.selfClosing ? undefined : event.localName
    } else if (
      event.kind === 'close' &&
      (event.localName === 'rowBreaks' || event.localName === 'colBreaks')
    ) {
      inside = undefined
    } else if (event.kind === 'open' && event.localName === 'brk' && inside !== undefined) {
      const id = numberAttribute(event.attributes, 'id') ?? 0
      if (inside === 'rowBreaks') rowIds.push(id)
      else columnIds.push(id)
    }
  }
  rowIds.sort((a, b) => a - b)
  columnIds.sort((a, b) => a - b)
  return {
    rows: rowIds.map((id) => id + 1),
    columns: columnIds.map((id) => indexToColumn(id + 1)),
  }
}
