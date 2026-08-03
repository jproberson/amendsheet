import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readWorkbook } from './document.js'
import type { CellInput } from './cell-input.js'
import { patchSheet as patchSheetBytes } from './patch.js'
import { XlsxError } from './errors.js'
import { assertPatchedSheet } from '../testing/invariants.js'

/** A1 is a fixed target, so on a real file it can land on a cell the file has
 * locked into a shared formula, array or merge. Refusing that overwrite is
 * correct behaviour, so the edit is tolerated and moved to a fresh cell. */
const isRefusal = (error: unknown) =>
  error instanceof XlsxError && error.code === 'unwritable-value'

const encode = (text: string) => new TextEncoder().encode(text)
const decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes)
const patchSheet = (
  source: string,
  edits: ReadonlyMap<string, CellInput>,
  date1904: boolean,
  sharedStrings?: ReadonlyMap<string, number>,
  styleOverrides?: ReadonlyMap<string, number>,
) => decode(patchSheetBytes(encode(source), edits, date1904, sharedStrings, styleOverrides))

const sheet = (rows: string) => `<worksheet><sheetData>${rows}</sheetData></worksheet>`

test('a cell added to a self closing row lands inside that row', () => {
  const patched = patchSheet(sheet('<row r="1"/>'), new Map<string, CellInput>([['A1', 1]]), false)

  assertPatchedSheet(patched, 'self closing row')
  assert.match(patched, /<row r="1"><c r="A1"><v>1<\/v><\/c><\/row>/)
})

test('setting a cell in a row whose cells omit references keeps the others', () => {
  const source = sheet('<row r="1"><c><v>10</v></c></row><row r="2"><c><v>20</v></c></row>')

  const patched = patchSheet(source, new Map<string, CellInput>([['B2', 99]]), false)

  assertPatchedSheet(patched, 'reference-less cells')
  assert.match(patched, /<v>20<\/v>/, 'the existing A2 value was destroyed')
  assert.match(patched, /<c r="B2"><v>99<\/v><\/c>/)
})

test('a reference given in any spelling is written in canonical form', () => {
  const source = sheet('<row r="1"><c r="A1"><v>1</v></c></row>')

  const patched = patchSheet(source, new Map<string, CellInput>([['$a$1', 5]]), false)

  assert.equal(patched.includes('$a$1'), false, 'the raw spelling reached the file')
  assert.match(patched, /<c r="A1"><v>5<\/v><\/c>/)
})

test('editing every real fixture leaves the sheet well formed', async () => {
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  const broken: string[] = []

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const workbook = readWorkbook(bytes)

    const target = workbook.sheets[0]
    if (target === undefined) continue

    const written: string[] = []
    for (const [reference, value] of [
      ['A1', 'edited'],
      ['ZZ900', 'far away'],
    ] as const) {
      try {
        target.set(reference, value)
        written.push(reference)
      } catch (error) {
        if (!isRefusal(error)) throw error
      }
    }

    try {
      const reopened = readWorkbook(workbook.toBytes())
      const edited = reopened.sheets[0]
      for (const reference of written) {
        assert.equal(
          edited?.cell(reference)?.value.kind,
          'text',
          `${file}: ${reference} did not read back`,
        )
      }
    } catch (error) {
      broken.push(`${file}: ${error instanceof Error ? error.message : String(error)}`)
    }
  }

  assert.deepEqual(broken, [])
})

test('editing a real fixture changes only the sheet and the tables it appends to', async () => {
  const { readContainer } = await import('./container.js')
  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  const unexpected: string[] = []

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(`fixtures/real/${file}`))
    const before = readContainer(bytes)

    const workbook = readWorkbook(bytes)
    const target = workbook.sheets[0]
    if (target === undefined) continue
    try {
      target.set('A1', 'edited')
    } catch (error) {
      if (!isRefusal(error)) throw error
      target.set('ZZ900', 'edited')
    }

    const after = readContainer(workbook.toBytes())
    const worksheets: string[] = []

    // Deliberate: the string and style tables grow, and a stale calculation
    // chain is dropped along with its content type override and the
    // relationship naming it.
    const expectedToChange = new Set([
      'xl/sharedStrings.xml',
      'xl/styles.xml',
      'xl/calcChain.xml',
      'xl/_rels/workbook.xml.rels',
      '[Content_Types].xml',
    ])

    for (const [path, content] of before.parts) {
      if (expectedToChange.has(path)) continue

      const other = after.parts.get(path)
      if (other === undefined) {
        unexpected.push(`${file}: lost ${path}`)
        continue
      }
      if (Buffer.compare(Buffer.from(content), Buffer.from(other)) === 0) continue

      if (path.toLowerCase().startsWith('xl/worksheets/')) {
        worksheets.push(path)
        continue
      }
      unexpected.push(`${file}: changed ${path}`)
    }

    if (worksheets.length > 1) {
      unexpected.push(`${file}: changed more than one sheet: ${worksheets.join(', ')}`)
    }
  }

  assert.deepEqual(unexpected, [])
})

const PREFIXED =
  '<x:worksheet><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>'

test('an edit to a prefixed document keeps the prefix', () => {
  const patched = patchSheet(PREFIXED, new Map<string, CellInput>([['A1', 2]]), false)

  assertPatchedSheet(patched, 'prefixed replace')
  assert.match(patched, /<x:c r="A1"><x:v>2<\/x:v><\/x:c>/)
})

test('a cell added to a prefixed document keeps the prefix', () => {
  const patched = patchSheet(PREFIXED, new Map<string, CellInput>([['B1', 5]]), false)

  assertPatchedSheet(patched, 'prefixed insert')
  assert.match(patched, /<x:c r="B1"><x:v>5<\/x:v><\/x:c>/)
})

test('a row added to a prefixed document keeps the prefix', () => {
  const patched = patchSheet(PREFIXED, new Map<string, CellInput>([['A3', 7]]), false)

  assertPatchedSheet(patched, 'prefixed new row')
  assert.match(patched, /<x:row r="3"><x:c r="A3"><x:v>7<\/x:v><\/x:c><\/x:row>/)
})

test('two cells added to one self closing row stay in a single row', () => {
  const patched = patchSheet(
    sheet('<row r="1"/>'),
    new Map<string, CellInput>([
      ['A1', 1],
      ['B1', 2],
    ]),
    false,
  )

  assertPatchedSheet(patched, 'two cells, one self closing row')
  assert.match(patched, /<row r="1"><c r="A1"><v>1<\/v><\/c><c r="B1"><v>2<\/v><\/c><\/row>/)
})
