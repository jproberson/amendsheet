import { XlsxError } from './errors.js'

const LAST_COLUMN = 16384
const LAST_ROW = 1048576

const LETTER_A = 'A'.charCodeAt(0)

export interface CellAddress {
  /** One based, matching how references are written. */
  readonly row: number
  readonly column: number
}

const isLetter = (character: string) =>
  (character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z')

const isDigit = (character: string) => character >= '0' && character <= '9'

/** Bijective base 26: there is no zero digit, so A-Z then AA, not A-Z then BA. */
export function columnToIndex(letters: string): number {
  let index = 0
  for (const letter of letters.toUpperCase()) {
    index = index * 26 + (letter.charCodeAt(0) - LETTER_A + 1)
  }
  return index
}

export function indexToColumn(index: number): string {
  if (index < 1 || index > LAST_COLUMN) {
    throw new XlsxError('bad-reference', `Column ${index} is outside the sheet`, {
      reference: String(index),
    })
  }

  let letters = ''
  let remaining = index
  while (remaining > 0) {
    const digit = (remaining - 1) % 26
    letters = String.fromCharCode(LETTER_A + digit) + letters
    remaining = (remaining - 1 - digit) / 26
  }
  return letters
}

/** Scanned rather than matched: this runs once per cell. */
export function parseReference(reference: string): CellAddress {
  let index = 0
  if (reference.charAt(index) === '$') index++

  const lettersStart = index
  while (index < reference.length && isLetter(reference.charAt(index))) index++
  const letters = reference.slice(lettersStart, index)

  if (reference.charAt(index) === '$') index++

  const digitsStart = index
  while (index < reference.length && isDigit(reference.charAt(index))) index++
  const digits = reference.slice(digitsStart, index)

  if (letters === '' || digits === '' || index !== reference.length) {
    throw new XlsxError('bad-reference', `"${reference}" is not a cell reference`, { reference })
  }

  return { row: Number(digits), column: columnToIndex(letters) }
}

/**
 * Reading stays lenient because real files contain references a sheet cannot
 * really hold, such as row zero. Writing does not: a reference the caller
 * supplies has to be one Excel will accept.
 */
export function parseWritableReference(reference: string): CellAddress {
  const address = parseReference(reference)
  const { row, column } = address
  if (row < 1 || row > LAST_ROW || column < 1 || column > LAST_COLUMN) {
    throw new XlsxError('bad-reference', `"${reference}" is outside the sheet`, { reference })
  }
  return address
}

export function formatReference(address: CellAddress): string {
  return `${indexToColumn(address.column)}${address.row}`
}

/**
 * The canonical spelling, or undefined when no column letter names it. Reading
 * is lenient, so a file may hold an address that cannot be spelled back, and
 * one of those must not take the lookups around it down with it. The row is
 * left alone: files really do contain row zero, and it formats fine.
 */
export function canonicalReference(address: CellAddress): string | undefined {
  if (address.column < 1 || address.column > LAST_COLUMN) return undefined
  return formatReference(address)
}
