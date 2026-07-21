import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import {
  assertOnlyTheSheetChanged,
  assertPatchedSheet,
  assertWellFormed,
} from '../testing/invariants.js'
import { dateToSerial } from './date.js'
import { type WriteOptions, readWorkbook } from './document.js'
import { XlsxError } from './errors.js'
import type { CellInput } from './patch.js'
import { indexToColumn } from './reference.js'

/**
 * Hand-written tests only cover failure modes somebody already imagined, which
 * is the wrong tool for a format this full of surprises. These generate edits
 * against the real files and assert what must hold for any of them.
 *
 * The seed is fixed so a failure reproduces, and a failing case is shrunk to
 * the smallest set of edits that still fails.
 */

/** Deterministic generator: the same seed always produces the same edits. */
function randomSource(seed: number) {
  let state = seed >>> 0
  return () => {
    state = (state * 1664525 + 1013904223) >>> 0
    return state / 0x100000000
  }
}

type Edit = readonly [string, CellInput, WriteOptions?]

function makeEdits(next: () => number, count: number): Edit[] {
  const values: CellInput[] = [
    0,
    -1,
    42.5,
    'text',
    '  padded  ',
    'quote " and & and <',
    true,
    false,
    null,
    new Date(2024, 5, 15),
    // After 1904, so it is representable under either epoch.
    new Date(1990, 11, 31),
    { formula: 'A1+1' },
    // Refused values, so the properties see edits that were turned down as well
    // as edits that went through. Without them a refusal is rare enough in the
    // corpus that "a refusal changes nothing" is asserted about almost nothing.
    Number.NaN,
    'a\u0000b',
    new Date(1800, 0, 1),
  ]

  // A format is refused after the value itself has been checked, which is the
  // only way to reach a refusal that arrives once the edit is already in hand.
  const formats: Array<WriteOptions | undefined> = [
    undefined,
    undefined,
    undefined,
    { numberFormat: '0.00' },
    { numberFormat: 'yyyy-mm-dd' },
    { numberFormat: `bad\u0000format` },
  ]

  const edits: Edit[] = []
  const used = new Set<string>()

  for (let index = 0; index < count; index++) {
    const column = indexToColumn(1 + Math.floor(next() * 30))
    const row = 1 + Math.floor(next() * 40)
    const reference = `${column}${row}`
    if (used.has(reference)) continue
    used.add(reference)

    const value = values[Math.floor(next() * values.length)]
    const options = formats[Math.floor(next() * formats.length)]
    edits.push([reference, value === undefined ? 0 : value, options])
  }

  return edits
}

/**
 * The library refuses a few edits on purpose, and `set()` is the only place it
 * may do so. Skipping a refusal anywhere else would let a property that never
 * ran report green.
 */
function isRefusal(error: unknown): boolean {
  return error instanceof XlsxError && error.code === 'unwritable-value'
}

interface Applied {
  readonly written: Uint8Array
  /** The edits `set()` accepted. A refused one leaves nothing behind. */
  readonly edits: Edit[]
}

function applyEdits(bytes: Uint8Array, edits: readonly Edit[]): Applied {
  const workbook = readWorkbook(bytes)
  const sheet = workbook.sheets[0]
  assert.ok(sheet !== undefined, 'fixture has no sheets')

  const accepted: Edit[] = []
  for (const [reference, value, options] of edits) {
    try {
      sheet.set(reference, value, options)
    } catch (error) {
      if (!isRefusal(error)) throw error
      continue
    }
    accepted.push([reference, value, options])
  }

  // Nothing is caught around toBytes(): once set() has accepted an edit, saving
  // it has to work.
  return { written: workbook.toBytes(), edits: accepted }
}

/** Reduces a failing case to the fewest edits that still fail. */
function shrink(
  bytes: Uint8Array,
  edits: readonly Edit[],
  check: (out: Uint8Array) => void,
): Edit[] {
  let smallest = [...edits]

  for (let index = 0; index < smallest.length; ) {
    const without = smallest.filter((_unused, position) => position !== index)
    let stillFails = false
    try {
      check(applyEdits(bytes, without).written)
    } catch {
      stillFails = true
    }
    if (stillFails && without.length > 0) smallest = without
    else index++
  }

  return smallest
}

async function fixtures(): Promise<string[]> {
  return (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
}

test('any set of edits leaves a sheet a valid sheet', async () => {
  const files = await fixtures()
  const next = randomSource(20260720)

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const edits = makeEdits(next, 12)

    const check = (out: Uint8Array) => {
      const reopened = readWorkbook(out)
      for (const sheet of reopened.sheets) {
        // Reading every cell proves the sheet still parses end to end.
        for (const _cell of sheet.cells()) void _cell
      }
    }

    try {
      check(applyEdits(bytes, edits).written)
    } catch (error) {
      const smallest = shrink(bytes, edits, check)
      assert.fail(
        `${file} broke with ${JSON.stringify(smallest)}: ` +
          `${error instanceof Error ? error.message : String(error)}`,
      )
    }
  }
})

test('every edited cell reads back as it was written', async () => {
  const files = await fixtures()
  const next = randomSource(11223344)

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const { written, edits } = applyEdits(bytes, makeEdits(next, 10))
    const workbook = readWorkbook(written)
    const sheet = workbook.sheets[0]

    for (const [reference, value] of edits) {
      const cell = sheet?.cell(reference)
      if (value === null) {
        assert.equal(cell?.value.kind, 'empty', `${file} ${reference}`)
        continue
      }
      if (typeof value === 'number') {
        // A number written into a date formatted cell keeps that format, so it
        // reads back as the date that serial means.
        if (cell?.value.kind === 'date') {
          assert.equal(cell.value.serial, value, `${file} ${reference}`)
        } else {
          assert.deepEqual(cell?.value, { kind: 'number', value }, `${file} ${reference}`)
        }
      } else if (typeof value === 'string') {
        assert.deepEqual(cell?.value, { kind: 'text', value }, `${file} ${reference}`)
      } else if (typeof value === 'boolean') {
        assert.deepEqual(cell?.value, { kind: 'boolean', value }, `${file} ${reference}`)
      } else if (value instanceof Date) {
        // A date is stored as a serial, and an asked-for format decides whether
        // it displays as one. Either way the serial has to be the same instant.
        const stored =
          cell?.value.kind === 'date'
            ? cell.value.serial
            : cell?.value.kind === 'number'
              ? cell.value.value
              : undefined
        assert.equal(stored, dateToSerial(value, workbook.date1904), `${file} ${reference}`)
      } else {
        // A formula carries no computed result, so only the expression is kept.
        assert.equal(cell?.formula, value.formula, `${file} ${reference}`)
      }
    }
  }
})

test('cells nobody edited are left alone', async () => {
  const files = await fixtures()
  const next = randomSource(55667788)

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    // Bounded on purpose: holding every cell of every file at once is enough
    // allocation to exhaust the heap, and a sample catches the same damage.
    const sample = 400
    const original = new Map<string, string>()
    for (const cell of readWorkbook(bytes).sheets[0]?.cells() ?? []) {
      if (original.size >= sample) break
      original.set(cell.reference, JSON.stringify(cell.value))
    }

    // A refused edit is not in `edits`, so the cell it aimed at is one of the
    // ones that has to come back unchanged.
    const { written, edits } = applyEdits(bytes, makeEdits(next, 8))
    const touched = new Set(edits.map(([reference]) => reference))

    // Looked up by reference rather than by position, since inserting a cell
    // shifts what the first N of the sheet are.
    const edited = readWorkbook(written).sheets[0]
    for (const [reference, was] of original) {
      if (touched.has(reference)) continue
      const now = edited?.cell(reference)
      assert.ok(now !== undefined, `${file} ${reference} disappeared`)
      assert.equal(JSON.stringify(now.value), was, `${file} ${reference} changed`)
    }
  }
})

test('no edit disturbs a part it has no business touching', async () => {
  const files = await fixtures()
  const next = randomSource(99001122)

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    assertOnlyTheSheetChanged(bytes, applyEdits(bytes, makeEdits(next, 6)).written, file)
  }
})

test('the order edits are applied in does not change what the sheet holds', async () => {
  const files = await fixtures()
  const next = randomSource(31415926)

  // The bytes may differ, since a new string lands wherever the table had room,
  // but what the sheet reports must not depend on the order.
  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const edits = makeEdits(next, 6)
    if (edits.length < 2) continue

    const forwards = applyEdits(bytes, edits)
    const backwards = applyEdits(bytes, [...edits].reverse())

    // Refusing an edit is order independent too.
    assert.deepEqual(
      forwards.edits.map(([reference]) => reference).sort(),
      backwards.edits.map(([reference]) => reference).sort(),
      `${file}: a different set of edits was refused`,
    )

    const one = readWorkbook(forwards.written).sheets[0]
    const other = readWorkbook(backwards.written).sheets[0]

    for (const [reference] of forwards.edits) {
      assert.deepEqual(
        one?.cell(reference)?.value,
        other?.cell(reference)?.value,
        `${file} ${reference} depends on the order edits were made`,
      )
    }

    let counted = 0
    for (const _cell of one?.cells() ?? []) counted++
    let alsoCounted = 0
    for (const _cell of other?.cells() ?? []) alsoCounted++
    assert.equal(counted, alsoCounted, `${file}: a different number of cells survived`)
  }
})

test('every patched sheet keeps its structure and its references', async () => {
  const files = await fixtures()
  const next = randomSource(27182818)
  const { readContainer } = await import('./container.js')

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const { written } = applyEdits(bytes, makeEdits(next, 10))

    for (const [path, content] of readContainer(written).parts) {
      const text = new TextDecoder().decode(content)
      if (path.toLowerCase().startsWith('xl/worksheets/') && path.endsWith('.xml')) {
        assertPatchedSheet(text, `${file} ${path}`)
      } else if (path.endsWith('.xml') || path.endsWith('.rels')) {
        assertWellFormed(text, `${file} ${path}`)
      }
    }
  }
})

/**
 * The other properties all read back from the written bytes, so a decision the
 * write path makes and the read path does not know about is invisible to them.
 * This one holds the live workbook against the file it produces.
 */
test('an edited workbook reads the same before and after it is written', async () => {
  const files = await fixtures()
  const next = randomSource(66778899)

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const edits = makeEdits(next, 10)

    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    assert.ok(sheet !== undefined, 'fixture has no sheets')

    const original = readWorkbook(bytes).sheets[0]
    const applied: Edit[] = []
    const refused: string[] = []
    for (const [reference, value, options] of edits) {
      try {
        sheet.set(reference, value, options)
      } catch (error) {
        if (!isRefusal(error)) throw error
        refused.push(reference)
        continue
      }
      applied.push([reference, value, options])
    }

    const reopened = readWorkbook(workbook.toBytes()).sheets[0]

    // A refusal has to leave no trace. Skipping these is what let an edit that
    // was rejected at the call still reach the file it was rejected from.
    for (const reference of refused) {
      const untouched = original?.cell(reference)
      assert.deepEqual(
        sheet.cell(reference)?.value,
        untouched?.value,
        `${file} ${reference} refused`,
      )
      assert.deepEqual(
        reopened?.cell(reference)?.value,
        untouched?.value,
        `${file} ${reference} refused, written`,
      )
    }

    for (const [reference] of applied) {
      const before = sheet.cell(reference)
      const after = reopened?.cell(reference)

      assert.deepEqual(before?.value, after?.value, `${file} ${reference} value`)
      assert.equal(before?.formula, after?.formula, `${file} ${reference} formula`)
      assert.equal(before?.numberFormat, after?.numberFormat, `${file} ${reference} number format`)
    }
  }
})
