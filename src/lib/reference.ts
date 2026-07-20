import { XlsxError } from './errors.js'

/** Sheets stop at column XFD; anything beyond it cannot be addressed. */
const LAST_COLUMN = 16384

const LETTER_A = 'A'.charCodeAt(0)

export interface CellAddress {
  /** One based, matching how references are written. */
  readonly row: number
  readonly column: number
}

const isLetter = (character: string) =>
  (character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z')

const isDigit = (character: string) => character >= '0' && character <= '9'

/**
 * Column letters are bijective base 26: A-Z, then AA-AZ, with no zero digit.
 * Treating them as ordinary base 26 puts every column past Z in the wrong place.
 */
export function columnToIndex(letters: string): number {
  let index = 0
  for (const letter of letters.toUpperCase()) {
    index = index * 26 + (letter.charCodeAt(0) - LETTER_A + 1)
  }
  return index
}

export function indexToColumn(index: number): string {
  if (index < 1 || index > LAST_COLUMN) {
    throw new XlsxError(`Column ${index} is outside the sheet`)
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

/** Scanned rather than matched with a regular expression: this runs once per cell. */
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
    throw new XlsxError(`"${reference}" is not a cell reference`)
  }

  return { row: Number(digits), column: columnToIndex(letters) }
}

export function formatReference(address: CellAddress): string {
  return `${indexToColumn(address.column)}${address.row}`
}
