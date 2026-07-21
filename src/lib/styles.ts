import { readXml } from './xml.js'

/** Implied by every workbook rather than written to the file. */
const BUILT_IN_FORMATS = new Map<number, string>([
  [0, 'General'],
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
])

/** The id of a built in format matching this code, when one does. */
export function builtInFormatId(code: string): number | undefined {
  for (const [id, builtIn] of BUILT_IN_FORMATS) {
    if (builtIn === code) return id
  }
  return undefined
}

export interface Styles {
  /** Custom format codes, keyed by the id cell formats reference. */
  readonly numberFormats: ReadonlyMap<number, string>
  /** Number format id per cell format, indexed by a cell's `s` attribute. */
  readonly cellFormats: readonly number[]
}

export function readStyles(xml: string): Styles {
  const numberFormats = new Map<number, string>()
  const cellFormats: number[] = []
  let inCellFormats = false
  let inNumberFormats = false

  for (const event of readXml(xml)) {
    if (event.kind === 'close') {
      if (event.localName === 'cellXfs') inCellFormats = false
      if (event.localName === 'numFmts') inNumberFormats = false
      continue
    }
    if (event.kind !== 'open') continue

    if (event.localName === 'numFmts') {
      inNumberFormats = !event.selfClosing
      continue
    }
    // dxfs carry their own numFmt elements, which would overwrite the real ones.
    if (event.localName === 'numFmt' && inNumberFormats) {
      const id = Number(event.attributes.get('numFmtId'))
      const code = event.attributes.get('formatCode')
      if (!Number.isNaN(id) && code !== undefined) numberFormats.set(id, code)
      continue
    }
    // cellStyleXfs holds the same element name, so only read inside cellXfs.
    if (event.localName === 'cellXfs') {
      inCellFormats = !event.selfClosing
      continue
    }
    if (event.localName === 'xf' && inCellFormats) {
      cellFormats.push(Number(event.attributes.get('numFmtId') ?? 0))
    }
  }

  return { numberFormats, cellFormats }
}

export function numberFormatOf(styles: Styles, styleIndex: number | undefined): string | undefined {
  if (styleIndex === undefined) return undefined

  const formatId = styles.cellFormats[styleIndex]
  if (formatId === undefined) return undefined

  return styles.numberFormats.get(formatId) ?? BUILT_IN_FORMATS.get(formatId)
}

/**
 * The parts of a format code a number can display with, stripped of the text
 * that only looks like tokens. A code has up to four sections — positive,
 * negative, zero, then text — and a numeric cell never displays with the
 * fourth, so `#,##0;-#,##0;0;"due "mmm` is not a date format. Without the
 * stripping, the letters in `"day"0.0` read as date tokens.
 */
function numericSections(code: string): string {
  let stripped = ''
  let index = 0
  let section = 1

  while (index < code.length) {
    const character = code.charAt(index)

    if (character === ';') {
      section++
      if (section > 3) break
      stripped += ';'
      index++
      continue
    }
    if (character === '"') {
      index++
      while (index < code.length && code.charAt(index) !== '"') index++
      index++
      continue
    }
    // `\` escapes the next character, `*` repeats it to fill the column and `_`
    // skips a width equal to it. All three make it a literal, not a token.
    if (character === '\\' || character === '*' || character === '_') {
      index += 2
      continue
    }
    if (character === '[') {
      const end = code.indexOf(']', index)
      // [h] is elapsed time; [Red] and conditions are not.
      const inside = end === -1 ? '' : code.slice(index + 1, end)
      if (/^[hms]+$/i.test(inside)) stripped += inside
      index = end === -1 ? code.length : end + 1
      continue
    }

    stripped += character
    index++
  }

  return stripped
}

const DATE_TOKEN = /[ymdhs]/i

/** `[h]`, `[m]` and `[s]` mark elapsed time, which is a duration rather than a date. */
const ELAPSED_TIME = /\[[hms]+\]/i

/** Whether a format code shows a date, rather than a number or a duration. */
export function isDateFormatCode(code: string | undefined): boolean {
  if (code === undefined || code === 'General') return false
  if (ELAPSED_TIME.test(code)) return false

  return DATE_TOKEN.test(numericSections(code))
}

export function isDateFormat(styles: Styles, styleIndex: number | undefined): boolean {
  return isDateFormatCode(numberFormatOf(styles, styleIndex))
}
