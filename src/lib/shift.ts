import { LAST_COLUMN, LAST_ROW, columnToIndex, indexToColumn } from './reference.js'

export interface ShiftSpec {
  readonly axis: 'row' | 'column'
  /** One-based line the insert or delete happens at. */
  readonly at: number
  /** Positive to insert lines, negative to delete them. */
  readonly delta: number
  /** The sheet whose rows or columns move; a reference targeting it shifts. */
  readonly editedSheet: string
  /** True when the formula being rewritten lives on the edited sheet. */
  readonly onCurrentSheet: boolean
}

interface Cell {
  readonly columnDollar: string
  readonly letters: string
  readonly rowDollar: string
  readonly digits: string
}

// How a value is shifted: a lone cell dies inside a deletion, while the two ends
// of a range clamp to the surviving edge instead so the range only shrinks.
type Mode = 'cell' | 'low' | 'high'

const isLetter = (character: string | undefined) =>
  character !== undefined &&
  ((character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z'))

const isDigit = (character: string | undefined) =>
  character !== undefined && character >= '0' && character <= '9'

// A reference that runs into one of these is part of a longer identifier (a
// defined name like A5B), not a reference.
const isNamePart = (character: string | undefined) =>
  isLetter(character) || isDigit(character) || character === '_' || character === '.'

const normalizeSheet = (raw: string): string => {
  const unquoted =
    raw.startsWith("'") && raw.endsWith("'") ? raw.slice(1, -1).replaceAll("''", "'") : raw
  return unquoted.toLowerCase()
}

// Consumes a quoted run starting at `open`, treating a doubled quote as an escape.
// Returns the index just past the closing quote, or the string end if unclosed.
const skipQuoted = (text: string, open: number, quote: string): number => {
  let index = open + 1
  while (index < text.length) {
    if (text[index] === quote) {
      if (text[index + 1] === quote) {
        index += 2
        continue
      }
      return index + 1
    }
    index += 1
  }
  return index
}

const NAME = /[A-Za-z_\\][A-Za-z0-9_.\\]*/y

const matchName = (text: string, at: number): number | undefined => {
  NAME.lastIndex = at
  return NAME.exec(text) === null ? undefined : NAME.lastIndex
}

// Scans a run while `kind` holds and it stays within `limit` characters.
const runOf = (
  text: string,
  from: number,
  kind: (character: string | undefined) => boolean,
  limit: number,
): number => {
  let index = from
  while (index - from < limit && kind(text[index])) index += 1
  return index
}

const matchCell = (text: string, at: number): { cell: Cell; end: number } | undefined => {
  let index = at
  const columnDollar = text[index] === '$' ? '$' : ''
  index += columnDollar.length
  const lettersEnd = runOf(text, index, isLetter, 3)
  const letters = text.slice(index, lettersEnd)
  if (letters === '') return undefined
  index = lettersEnd
  const rowDollar = text[index] === '$' ? '$' : ''
  index += rowDollar.length
  const digitsEnd = runOf(text, index, isDigit, 7)
  const digits = text.slice(index, digitsEnd)
  // A run of digits capped at seven still followed by one is not a real row.
  if (digits === '' || isNamePart(text[digitsEnd])) return undefined
  return { cell: { columnDollar, letters, rowDollar, digits }, end: digitsEnd }
}

const matchWhole = (
  text: string,
  at: number,
): { dollar: string; body: string; numeric: boolean; end: number } | undefined => {
  let index = at
  const dollar = text[index] === '$' ? '$' : ''
  index += dollar.length
  const numeric = isDigit(text[index])
  const end = runOf(text, index, numeric ? isDigit : isLetter, numeric ? 7 : 3)
  const body = text.slice(index, end)
  if (body === '' || isNamePart(text[end])) return undefined
  return { dollar, body, numeric, end }
}

/** Where a line lands after the edit, or undefined when the edit removes it. */
const shiftAxisValue = (
  value: number,
  spec: ShiftSpec,
  max: number,
  mode: Mode,
): number | undefined => {
  if (value < spec.at) return value
  if (spec.delta > 0) {
    const moved = value + spec.delta
    return moved > max ? undefined : moved
  }
  const count = -spec.delta
  if (value < spec.at + count) {
    if (mode === 'cell') return undefined
    return mode === 'low' ? spec.at : spec.at - 1
  }
  return value + spec.delta
}

const maxOf = (spec: ShiftSpec): number => (spec.axis === 'row' ? LAST_ROW : LAST_COLUMN)

const rawCell = (cell: Cell, spec: ShiftSpec): number =>
  spec.axis === 'row' ? Number(cell.digits) : columnToIndex(cell.letters)

const rebuildCell = (cell: Cell, spec: ShiftSpec, value: number): string =>
  spec.axis === 'row'
    ? `${cell.columnDollar}${cell.letters}${cell.rowDollar}${value}`
    : `${cell.columnDollar}${indexToColumn(value)}${cell.rowDollar}${cell.digits}`

const REF = '#REF!'

const cellText = (cell: Cell, spec: ShiftSpec, mode: Mode): string => {
  const shifted = shiftAxisValue(rawCell(cell, spec), spec, maxOf(spec), mode)
  return shifted === undefined ? REF : rebuildCell(cell, spec, shifted)
}

const rangeText = (first: Cell, second: Cell, spec: ShiftSpec): string => {
  const max = maxOf(spec)
  const low = shiftAxisValue(rawCell(first, spec), spec, max, 'low')
  const high = shiftAxisValue(rawCell(second, spec), spec, max, 'high')
  if (low === undefined || high === undefined || low > high) return REF
  return `${rebuildCell(first, spec, low)}:${rebuildCell(second, spec, high)}`
}

// A whole-column (A:A) or whole-row (5:5) range end carries one axis and spans
// everything on the other, so it only moves when its axis is the edited one.
const wholeText = (
  low: { dollar: string; body: string; numeric: boolean },
  high: { dollar: string; body: string },
  spec: ShiftSpec,
): string => {
  const numeric = low.numeric
  if (numeric !== (spec.axis === 'row'))
    return `${low.dollar}${low.body}:${high.dollar}${high.body}`
  const value = (body: string) => (numeric ? Number(body) : columnToIndex(body))
  const spell = (line: number) => (numeric ? String(line) : indexToColumn(line))
  const lowLine = shiftAxisValue(value(low.body), spec, maxOf(spec), 'low')
  const highLine = shiftAxisValue(value(high.body), spec, maxOf(spec), 'high')
  if (lowLine === undefined || highLine === undefined || lowLine > highLine) return REF
  return `${low.dollar}${spell(lowLine)}:${high.dollar}${spell(highLine)}`
}

// Parses a reference or range at `at`, past any sheet qualifier. Returns its
// rewritten text and where it ends, or undefined when it is not a reference.
const shiftReference = (
  text: string,
  at: number,
  spec: ShiftSpec,
  targets: boolean,
): { text: string; end: number } | undefined => {
  const first = matchCell(text, at)
  if (first !== undefined) {
    if (text[first.end] !== ':') {
      if (!targets || text[first.end] === '(')
        return { text: text.slice(at, first.end), end: first.end }
      return { text: cellText(first.cell, spec, 'cell'), end: first.end }
    }
    const second = matchCell(text, first.end + 1)
    if (second === undefined) return { text: text.slice(at, first.end), end: first.end }
    if (!targets) return { text: text.slice(at, second.end), end: second.end }
    return { text: rangeText(first.cell, second.cell, spec), end: second.end }
  }

  const low = matchWhole(text, at)
  if (low === undefined || text[low.end] !== ':') return undefined
  const high = matchWhole(text, low.end + 1)
  if (high === undefined) return undefined
  if (!targets) return { text: text.slice(at, high.end), end: high.end }
  return { text: wholeText(low, high, spec), end: high.end }
}

/** Rewrites every reference in a formula for an inserted or deleted row or column. */
export function shiftFormula(formula: string, spec: ShiftSpec): string {
  const edited = normalizeSheet(spec.editedSheet)
  let out = ''
  let index = 0
  while (index < formula.length) {
    const character = formula[index]

    if (character === '"') {
      const end = skipQuoted(formula, index, '"')
      out += formula.slice(index, end)
      index = end
      continue
    }

    // A quoted or bare sheet name before a '!' qualifies the reference after it.
    let cursor = index
    let targets = spec.onCurrentSheet
    if (character === "'") {
      const close = skipQuoted(formula, index, "'")
      if (formula[close] === '!') {
        targets = normalizeSheet(formula.slice(index, close)) === edited
        cursor = close + 1
      } else {
        out += formula.slice(index, close)
        index = close
        continue
      }
    } else {
      const nameEnd = matchName(formula, index)
      if (nameEnd !== undefined && formula[nameEnd] === '!') {
        targets = normalizeSheet(formula.slice(index, nameEnd)) === edited
        cursor = nameEnd + 1
      }
    }

    const reference = shiftReference(formula, cursor, spec, targets)
    if (reference !== undefined) {
      out += formula.slice(index, cursor) + reference.text
      index = reference.end
      continue
    }

    // Not a reference. Consume any whole name so we do not rescan inside it.
    const nameEnd = matchName(formula, cursor)
    if (nameEnd !== undefined) {
      out += formula.slice(index, nameEnd)
      index = nameEnd
      continue
    }
    out += formula.slice(index, cursor + 1)
    index = cursor + 1
  }
  return out
}
