import assert from 'node:assert/strict'
import { test } from 'node:test'
import { createStylesSession } from './styles-session.js'
import {
  type Alignment,
  type BorderFormat,
  type CellProtection,
  type DateStyle,
  type FillFormat,
  type FontFormat,
  ensureAlignmentStyle,
  ensureBorderStyle,
  ensureDateStyle,
  ensureFillStyle,
  ensureFontStyle,
  ensureNumberFormat,
  ensureProtectionStyle,
  readFormatting,
} from './styles-writer.js'
import { numberFormatOf, readStyles } from './styles.js'

// One set()/format() call: a number format or date, then any of font, fill,
// border, alignment, protection, each landing on the one before — the order
// resolveStyle threads them in.
interface Op {
  readonly base: 'new' | 'prev' | number
  readonly numberFormat?: string
  readonly date?: boolean
  readonly font?: FontFormat
  readonly fill?: FillFormat
  readonly border?: BorderFormat
  readonly alignment?: Alignment
  readonly protection?: CellProtection
}

const LOCATION = {}

function applyOracle(
  xml: string,
  base: number | undefined,
  op: Op,
): { xml: string; index: number } {
  let current = base
  let working = xml
  const step = (result: DateStyle): void => {
    working = result.xml
    current = result.index
  }
  if (op.numberFormat !== undefined)
    step(ensureNumberFormat(working, current, op.numberFormat, LOCATION))
  else if (op.date === true) step(ensureDateStyle(working, current))
  if (op.font !== undefined) step(ensureFontStyle(working, current, op.font, LOCATION))
  if (op.fill !== undefined) step(ensureFillStyle(working, current, op.fill, LOCATION))
  if (op.border !== undefined) step(ensureBorderStyle(working, current, op.border, LOCATION))
  if (op.alignment !== undefined)
    step(ensureAlignmentStyle(working, current, op.alignment, LOCATION))
  if (op.protection !== undefined) step(ensureProtectionStyle(working, current, op.protection))
  return { xml: working, index: current ?? 0 }
}

function applySession(
  session: ReturnType<typeof createStylesSession>,
  base: number | undefined,
  op: Op,
): number {
  let current = base
  if (op.numberFormat !== undefined)
    current = session.numberFormat(current, op.numberFormat, LOCATION)
  else if (op.date === true) current = session.dateStyle(current)
  if (op.font !== undefined) current = session.font(current, op.font, LOCATION)
  if (op.fill !== undefined) current = session.fill(current, op.fill, LOCATION)
  if (op.border !== undefined) current = session.border(current, op.border, LOCATION)
  if (op.alignment !== undefined) current = session.alignment(current, op.alignment, LOCATION)
  if (op.protection !== undefined) current = session.protection(current, op.protection)
  return current ?? 0
}

/** Applies the same ops through the oracle and the session, asserting the styles
 * come out byte-identical and every returned index matches, then that the read
 * accessors agree with a fresh read of the oracle's output. */
function checkIdentical(startXml: string, ops: readonly Op[]): void {
  let oracleXml = startXml
  const session = createStylesSession(startXml)
  let previous: number | undefined
  for (const op of ops) {
    const base = op.base === 'new' ? undefined : op.base === 'prev' ? previous : op.base
    const oracle = applyOracle(oracleXml, base, op)
    const index = session.transaction(() => applySession(session, base, op))
    assert.equal(index, oracle.index, `index diverged for ${JSON.stringify(op)}`)
    oracleXml = oracle.xml
    previous = oracle.index
  }

  assert.equal(session.serialize(), oracleXml, 'serialized styles diverged')

  const formatting = readFormatting(oracleXml)
  const styles = readStyles(oracleXml)
  const count = readStyles(oracleXml).cellFormats.length
  for (let index = 0; index < count; index++) {
    assert.deepEqual(session.formattingOf(index), formatting[index] ?? {}, `formatting ${index}`)
    assert.equal(
      session.numberFormatOf(index),
      numberFormatOf(styles, index),
      `numberFormat ${index}`,
    )
  }
}

const BLANK_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '</styleSheet>'

// A hand-written file with a namespace prefix, populated tables and whitespace
// between elements, so the session has to preserve bytes it does not touch.
const PREFIXED_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n' +
  '  <x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.000"/></x:numFmts>\n' +
  '  <x:fonts count="2">\n' +
  '    <x:font><x:sz val="11"/><x:name val="Calibri"/></x:font>\n' +
  '    <x:font><x:b/><x:sz val="12"/><x:color rgb="FF0000FF"/></x:font>\n' +
  '  </x:fonts>\n' +
  '  <x:fills count="3">\n' +
  '    <x:fill><x:patternFill patternType="none"/></x:fill>\n' +
  '    <x:fill><x:patternFill patternType="gray125"/></x:fill>\n' +
  '    <x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FFFF0000"/><x:bgColor indexed="64"/></x:patternFill></x:fill>\n' +
  '  </x:fills>\n' +
  '  <x:borders count="2">\n' +
  '    <x:border><x:left/><x:right/><x:top/><x:bottom/><x:diagonal/></x:border>\n' +
  '    <x:border><x:left style="thin"><x:color rgb="FF000000"/></x:left><x:right/><x:top/><x:bottom/><x:diagonal/></x:border>\n' +
  '  </x:borders>\n' +
  '  <x:cellXfs count="2">\n' +
  '    <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n' +
  '    <x:xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>\n' +
  '  </x:cellXfs>\n' +
  '</x:styleSheet>'

// A file whose numFmts already holds custom formats, so allocation has to skip
// the ids it uses and matching an existing code reuses it.
const NUMFMT_STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.0"/><numFmt numFmtId="166" formatCode="0.00000"/></numFmts>' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
  '</styleSheet>'

test('distinct number formats each allocate a new custom id', () => {
  const ops: Op[] = []
  for (let i = 0; i < 40; i++)
    ops.push({ base: 'new', numberFormat: `0.${'0'.repeat((i % 6) + 1)}"code${i}"` })
  checkIdentical(BLANK_STYLES, ops)
})

test('reused number formats collapse onto one cell format', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', numberFormat: '0.00' },
    { base: 'new', numberFormat: '0.00' },
    { base: 'new', numberFormat: '"$"#,##0.00' },
    { base: 'new', numberFormat: '0.00' },
    { base: 'new', numberFormat: '"$"#,##0.00' },
  ])
})

test('dates reuse and borrow a date-formatted cell format', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', date: true },
    { base: 'new', date: true },
    { base: 'prev', date: true },
    { base: 'new', numberFormat: 'yyyy-mm-dd' },
    { base: 'new', date: true },
  ])
})

test('font variants and merges onto an existing font', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', font: { bold: true } },
    { base: 'new', font: { italic: true } },
    { base: 'new', font: { underline: true } },
    { base: 'new', font: { underline: 'double' } },
    { base: 'new', font: { underline: 'singleAccounting' } },
    { base: 'new', font: { size: 18 } },
    { base: 'new', font: { color: '00FF00' } },
    { base: 'new', font: { color: { theme: 4, tint: -0.25 } } },
    { base: 'new', font: { name: 'Arial' } },
    { base: 'new', font: { strike: true, verticalAlign: 'superscript' } },
    { base: 1, font: { italic: true } },
    { base: 1, font: { bold: true, size: 22 } },
  ])
})

test('fills: solid, pattern and gradient', () => {
  const ops: Op[] = [
    { base: 'new', fill: { type: 'solid', color: 'FF0000' } },
    { base: 'new', fill: { type: 'solid', color: 'FF0000' } },
    {
      base: 'new',
      fill: { type: 'pattern', pattern: 'darkGrid', color: '112233', background: 'AABBCC' },
    },
    { base: 'new', fill: { type: 'pattern', pattern: 'gray125', color: '000000' } },
    { base: 'new', fill: { type: 'pattern', pattern: 'lightUp', color: '00FF00' } },
    {
      base: 'new',
      fill: {
        type: 'gradient',
        degree: 45,
        stops: [
          { position: 0, color: 'FFFF0000' },
          { position: 1, color: 'FF0000FF' },
        ],
      },
    },
    { base: 'new', fill: { type: 'gradient', stops: [{ position: 0, color: 'FF000000' }] } },
    { base: 3, fill: { type: 'solid', color: '0000FF' } },
  ]
  checkIdentical(BLANK_STYLES, ops)
})

test('borders: per side, all, diagonal, colours', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', border: { all: { style: 'thin' } } },
    { base: 'new', border: { top: { style: 'thick', color: 'FF0000' } } },
    { base: 'new', border: { left: { style: 'dashed' }, right: { style: 'dotted' } } },
    { base: 'new', border: { diagonal: { style: 'thin', up: true } } },
    { base: 'new', border: { diagonal: { style: 'medium', down: true, color: '0000FF' } } },
    { base: 'new', border: { all: { style: 'thin' }, bottom: { style: 'double' } } },
    { base: 4, border: { top: { style: 'hair' } } },
  ])
})

test('alignment: every attribute', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', alignment: { horizontal: 'center' } },
    { base: 'new', alignment: { vertical: 'top' } },
    { base: 'new', alignment: { wrapText: true } },
    { base: 'new', alignment: { textRotation: 45 } },
    { base: 'new', alignment: { textRotation: 255 } },
    { base: 'new', alignment: { indent: 3 } },
    { base: 1, alignment: { vertical: 'center', wrapText: true } },
    {
      base: 'new',
      alignment: {
        horizontal: 'right',
        vertical: 'bottom',
        wrapText: true,
        textRotation: 90,
        indent: 2,
      },
    },
  ])
})

test('protection', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', protection: { locked: false } },
    { base: 'new', protection: { hidden: true } },
    { base: 'new', protection: { locked: true, hidden: false } },
    { base: 1, protection: { hidden: true } },
  ])
})

test('a number format, font, fill, border and alignment chained in one op', () => {
  checkIdentical(BLANK_STYLES, [
    {
      base: 'new',
      numberFormat: '"$"#,##0.00',
      font: { bold: true, color: 'FFFFFF' },
      fill: { type: 'solid', color: '004488' },
      border: { all: { style: 'thin' } },
      alignment: { horizontal: 'center' },
      protection: { locked: false },
    },
    { base: 'prev', font: { italic: true } },
    { base: 'prev', numberFormat: '0.0' },
  ])
})

test('format() restyles an already styled cell', () => {
  checkIdentical(BLANK_STYLES, [
    { base: 'new', numberFormat: '0.00', font: { bold: true } },
    { base: 'prev', fill: { type: 'solid', color: 'FF0000' } },
    { base: 'prev', border: { all: { style: 'thin' } } },
    { base: 'prev', alignment: { wrapText: true } },
  ])
})

test('starting from a prefixed, populated, whitespaced styles.xml', () => {
  checkIdentical(PREFIXED_STYLES, [
    { base: 'new', numberFormat: '0.000' },
    { base: 'new', numberFormat: '0.0000' },
    { base: 1, font: { italic: true } },
    { base: 1, fill: { type: 'solid', color: '00FF00' } },
    { base: 'new', border: { all: { style: 'thick' } } },
    { base: 'new', alignment: { horizontal: 'center' } },
    { base: 1, numberFormat: '0.000' },
    { base: 1, date: true },
  ])
})

test('starting from populated numFmts skips used ids and reuses codes', () => {
  checkIdentical(NUMFMT_STYLES, [
    { base: 'new', numberFormat: '0.0' },
    { base: 'new', numberFormat: '0.00000' },
    { base: 'new', numberFormat: '0.000000' },
    { base: 'new', numberFormat: '0.0000000' },
    { base: 1, numberFormat: '0.0' },
    { base: 'new', numberFormat: '0.0' },
  ])
})

// Self-closing empty tables, which an odd file writes as `<fonts count="0"/>`.
// Each is seeded or opened on first use.
const SELF_CLOSING_STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="0"/><fonts count="0"/><fills count="0"/><borders count="0"/><cellXfs count="0"/>' +
  '</styleSheet>'

// Empty but paired tables, so the seed replaces the whole container rather than a
// self-closing tag.
const EMPTY_PRESENT_STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts></fonts><fills></fills><borders></borders><cellXfs></cellXfs>' +
  '</styleSheet>'

// No sub-tables at all, so every one is created at the point the oracle opens it.
const BARE_STYLES =
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>'

test('self-closing empty tables are seeded and opened byte-identically', () => {
  checkIdentical(SELF_CLOSING_STYLES, [
    { base: 'new', numberFormat: '0.000' },
    { base: 'new', numberFormat: '0.000' },
    { base: 'new', font: { bold: true } },
    { base: 'new', fill: { type: 'solid', color: 'FF0000' } },
    { base: 'new', border: { all: { style: 'thin' } } },
    { base: 'new', alignment: { horizontal: 'center' } },
    { base: 'new', date: true },
  ])
})

test('empty paired tables are replaced with a seeded container', () => {
  checkIdentical(EMPTY_PRESENT_STYLES, [
    { base: 'new', font: { italic: true } },
    { base: 'new', fill: { type: 'pattern', pattern: 'darkGrid', color: '112233' } },
    { base: 'new', border: { top: { style: 'thick' } } },
    { base: 'new', numberFormat: '0.00' },
  ])
})

test('absent tables are created where the oracle opens them', () => {
  checkIdentical(BARE_STYLES, [
    { base: 'new', font: { bold: true } },
    { base: 'new', fill: { type: 'solid', color: '00FF00' } },
    { base: 'new', border: { all: { style: 'thin' } } },
    { base: 'new', numberFormat: '0.000' },
    { base: 'new', alignment: { wrapText: true } },
    { base: 'new', date: true },
  ])
})

test('the number-format order does not matter when tables are absent', () => {
  checkIdentical(BARE_STYLES, [
    { base: 'new', numberFormat: '0.000' },
    { base: 'new', border: { all: { style: 'thin' } } },
    { base: 'new', font: { bold: true } },
    { base: 'new', fill: { type: 'solid', color: '00FF00' } },
  ])
})

test('a styleSheet with no place to create a table throws malformed', () => {
  const noRoot = createStylesSession('<other xmlns="http://x"></other>')
  noRoot.transaction(() => noRoot.numberFormat(undefined, '0.000000', LOCATION))
  assert.throws(() => noRoot.serialize(), /styleSheet/)

  const noClose = createStylesSession('<styleSheet xmlns="http://x">')
  noClose.transaction(() => noClose.font(undefined, { bold: true }, LOCATION))
  assert.throws(() => noClose.serialize(), /styleSheet/)
})

test('a throw mid-op rolls the session back to match the oracle', () => {
  const session = createStylesSession(BLANK_STYLES)
  let oracleXml = BLANK_STYLES

  // A first valid op, applied to both.
  const first: Op = { base: 'new', numberFormat: '0.00' }
  const oracleFirst = applyOracle(oracleXml, undefined, first)
  const sessionFirst = session.transaction(() => applySession(session, undefined, first))
  assert.equal(sessionFirst, oracleFirst.index)
  oracleXml = oracleFirst.xml

  // A op that adds a number format then throws on a bad font size. The oracle
  // never commits, so its xml is unchanged; the session must roll back too.
  const bad: Op = { base: 'new', numberFormat: '0.000', font: { size: -3 } }
  assert.throws(() => applyOracle(oracleXml, undefined, bad))
  assert.throws(() => session.transaction(() => applySession(session, undefined, bad)))
  assert.equal(session.serialize(), oracleXml, 'rolled-back styles diverged')

  // A following valid op must still line up.
  const third: Op = { base: 'new', numberFormat: '0.000' }
  const oracleThird = applyOracle(oracleXml, undefined, third)
  const sessionThird = session.transaction(() => applySession(session, undefined, third))
  assert.equal(sessionThird, oracleThird.index)
  assert.equal(session.serialize(), oracleThird.xml)
})

test('bad values throw the same error the oracle throws', () => {
  const session = createStylesSession(BLANK_STYLES)
  assert.throws(
    () => session.transaction(() => session.numberFormat(undefined, 'a b', LOCATION)),
    /cannot be written to xml/,
  )
  assert.throws(
    () => session.transaction(() => session.font(undefined, { size: 0 }, LOCATION)),
    /is not a positive number/,
  )
  assert.throws(
    () => session.transaction(() => session.font(undefined, { color: 'nothex' }, LOCATION)),
    /6 or 8 digit hex/,
  )
  assert.throws(
    () => session.transaction(() => session.alignment(undefined, { textRotation: 400 }, LOCATION)),
    /0–180 or 255/,
  )
})

// A small deterministic LCG so the randomised sequence is reproducible without
// Math.random, which the house rules ban.
function lcg(seed: number): () => number {
  let state = seed >>> 0
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state / 0x100000000
  }
}

test('a long randomised sequence stays byte-identical', () => {
  const random = lcg(0x51ed_5eed)
  const pick = <T>(items: readonly T[]): T => {
    const item = items[Math.floor(random() * items.length)]
    if (item === undefined) throw new Error('pick from an empty list')
    return item
  }
  const codes = ['0.0', '0.00', '"$"#,##0.00', 'yyyy-mm-dd', '0%', '@', '#,##0']
  const colors = ['FF0000', '00FF00', '0000FF', 'ABCDEF', '123456']
  const patterns = ['darkGrid', 'lightUp', 'gray0625'] as const
  const borderStyles = ['thin', 'medium', 'thick', 'dashed', 'double'] as const
  const horizontals = ['left', 'center', 'right', 'justify'] as const

  const ops: Op[] = []
  let known = 1
  for (let i = 0; i < 200; i++) {
    const base: Op['base'] = random() < 0.4 ? 'new' : Math.floor(random() * known)
    const op: {
      base: Op['base']
      numberFormat?: string
      date?: boolean
      font?: FontFormat
      fill?: FillFormat
      border?: BorderFormat
      alignment?: Alignment
      protection?: CellProtection
    } = { base }
    const kind = Math.floor(random() * 7)
    if (kind === 0) op.numberFormat = pick(codes)
    else if (kind === 1) op.date = true
    else if (kind === 2)
      op.font = { bold: random() < 0.5, size: 8 + Math.floor(random() * 16), color: pick(colors) }
    else if (kind === 3)
      op.fill =
        random() < 0.5
          ? { type: 'solid', color: pick(colors) }
          : { type: 'pattern', pattern: pick(patterns), color: pick(colors) }
    else if (kind === 4)
      op.border = {
        all: { style: pick(borderStyles) },
        top: { style: pick(borderStyles), color: pick(colors) },
      }
    else if (kind === 5) op.alignment = { horizontal: pick(horizontals), wrapText: random() < 0.5 }
    else op.protection = { locked: random() < 0.5, hidden: random() < 0.5 }
    ops.push(op)
    known += 1
  }
  checkIdentical(BLANK_STYLES, ops)
})
