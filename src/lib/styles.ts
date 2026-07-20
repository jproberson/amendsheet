import { readXml } from './xml.js'

/**
 * Number formats built into every workbook rather than written to the file.
 * Only the ones that matter for telling dates from numbers are listed.
 */
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

  for (const event of readXml(xml)) {
    if (event.kind === 'close') {
      if (event.name === 'cellXfs') inCellFormats = false
      continue
    }
    if (event.kind !== 'open') continue

    if (event.name === 'numFmt') {
      const id = Number(event.attributes.get('numFmtId'))
      const code = event.attributes.get('formatCode')
      if (!Number.isNaN(id) && code !== undefined) numberFormats.set(id, code)
      continue
    }
    // cellStyleXfs holds the same element name, so only read inside cellXfs.
    if (event.name === 'cellXfs') {
      inCellFormats = !event.selfClosing
      continue
    }
    if (event.name === 'xf' && inCellFormats) {
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
 * Strips the parts of a format code that are shown literally, so their letters
 * are not mistaken for date tokens: `"day"0.0` is a number, not a date.
 * Elapsed-time brackets such as `[h]` are kept, since those are times.
 */
function stripLiterals(code: string): string {
  let stripped = ''
  let index = 0

  while (index < code.length) {
    const character = code.charAt(index)

    if (character === '"') {
      index++
      while (index < code.length && code.charAt(index) !== '"') index++
      index++
      continue
    }
    if (character === '\\') {
      index += 2
      continue
    }
    if (character === '[') {
      const end = code.indexOf(']', index)
      const section = end === -1 ? '' : code.slice(index + 1, end)
      // [h], [m] and [s] are elapsed time; anything else is a colour or condition.
      if (/^[hms]+$/i.test(section)) stripped += section
      index = end === -1 ? code.length : end + 1
      continue
    }

    stripped += character
    index++
  }

  return stripped
}

const DATE_TOKEN = /[ymdhs]/i

export function isDateFormat(styles: Styles, styleIndex: number | undefined): boolean {
  const code = numberFormatOf(styles, styleIndex)
  if (code === undefined || code === 'General') return false

  return DATE_TOKEN.test(stripLiterals(code))
}
