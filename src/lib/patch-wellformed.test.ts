import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { readWorkbook } from './document.js'
import { type CellInput, patchSheet } from './patch.js'
import { readXml } from './xml.js'

/**
 * Splicing can produce output that contains the expected substring and is still
 * broken, so every patch result is checked for structure rather than fragments.
 */
function assertWellFormed(xml: string, what: string): void {
  const open: string[] = []
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && !event.selfClosing) open.push(event.name)
    if (event.kind === 'close') assert.equal(open.pop(), event.name, `${what}: mismatched close`)
  }
  assert.deepEqual(open, [], `${what}: unclosed elements`)
}

/** Every cell must sit inside a row, and every row inside sheetData. */
function assertCellsAreInRows(xml: string, what: string): void {
  let inRow = false
  let inData = false
  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.name === 'sheetData') inData = !event.selfClosing
      if (event.name === 'row') {
        assert.ok(inData, `${what}: row outside sheetData`)
        inRow = !event.selfClosing
      }
      if (event.name === 'c') assert.ok(inRow, `${what}: cell outside a row`)
    }
    if (event.kind === 'close') {
      if (event.name === 'row') inRow = false
      if (event.name === 'sheetData') inData = false
    }
  }
}

const sheet = (rows: string) => `<worksheet><sheetData>${rows}</sheetData></worksheet>`

test('a cell added to a self closing row lands inside that row', () => {
  const patched = patchSheet(sheet('<row r="1"/>'), new Map<string, CellInput>([['A1', 1]]), false)

  assertWellFormed(patched, 'self closing row')
  assertCellsAreInRows(patched, 'self closing row')
  assert.match(patched, /<row r="1"><c r="A1"><v>1<\/v><\/c><\/row>/)
})

test('setting a cell in a row whose cells omit references keeps the others', () => {
  const source = sheet('<row r="1"><c><v>10</v></c></row><row r="2"><c><v>20</v></c></row>')

  const patched = patchSheet(source, new Map<string, CellInput>([['B2', 99]]), false)

  assertWellFormed(patched, 'reference-less cells')
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

    target.set('A1', 'edited')
    target.set('ZZ900', 'far away')

    try {
      const reopened = readWorkbook(workbook.toBytes())
      const edited = reopened.sheets[0]
      assert.equal(edited?.cell('A1')?.value.kind, 'text', `${file}: A1 did not read back`)
      assert.equal(edited?.cell('ZZ900')?.value.kind, 'text', `${file}: ZZ900 did not read back`)
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
    target.set('A1', 'edited')

    const after = readContainer(workbook.toBytes())
    const worksheets: string[] = []

    for (const [path, content] of before.parts) {
      const other = after.parts.get(path)
      if (other === undefined) {
        unexpected.push(`${file}: lost ${path}`)
        continue
      }
      if (Buffer.compare(Buffer.from(content), Buffer.from(other)) === 0) continue

      if (path === 'xl/sharedStrings.xml' || path === 'xl/styles.xml') continue
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
