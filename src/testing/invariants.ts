import assert from 'node:assert/strict'
import { readContainer } from '../lib/container.js'
import { readXml } from '../lib/xml.js'

/**
 * Assertions every write has to satisfy. Tests that check for a substring pass
 * on output that contains the right fragment inside a broken document, which is
 * how three corrupting bugs shipped with green tests.
 *
 * Lives outside src/lib so it is not measured as library code.
 */

export function assertWellFormed(xml: string, what: string): void {
  const open: string[] = []
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && !event.selfClosing) open.push(event.name)
    if (event.kind === 'close') assert.equal(open.pop(), event.name, `${what}: mismatched close`)
  }
  assert.deepEqual(open, [], `${what}: unclosed elements`)
}

/** A cell outside a row, or a row outside sheetData, makes Excel repair the file. */
export function assertSheetShape(xml: string, what: string): void {
  let inRow = false
  let inData = false

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.localName === 'sheetData') inData = !event.selfClosing
      if (event.localName === 'row') {
        assert.ok(inData, `${what}: row outside sheetData`)
        inRow = !event.selfClosing
      }
      if (event.localName === 'c') assert.ok(inRow, `${what}: cell outside a row`)
    }
    if (event.kind === 'close') {
      if (event.localName === 'row') inRow = false
      if (event.localName === 'sheetData') inData = false
    }
  }
}

/** No cell reference may appear twice, and every one must be canonical. */
export function assertReferencesAreSane(xml: string, what: string): void {
  const seen = new Set<string>()
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'c') continue
    const reference = event.attributes.get('r')
    if (reference === undefined) continue

    assert.match(reference, /^[A-Z]+[0-9]+$/, `${what}: ${reference} is not a canonical reference`)
    assert.ok(!seen.has(reference), `${what}: ${reference} appears more than once`)
    seen.add(reference)
  }
}

/** Two rows with the same index make Excel repair the file. */
export function assertRowsAreUnique(xml: string, what: string): void {
  const seen = new Set<string>()
  for (const event of readXml(xml)) {
    if (event.kind !== 'open' || event.localName !== 'row') continue
    const index = event.attributes.get('r')
    if (index === undefined) continue
    assert.ok(!seen.has(index), `${what}: row ${index} appears more than once`)
    seen.add(index)
  }
}

export function assertPatchedSheet(xml: string, what: string): void {
  assertWellFormed(xml, what)
  assertSheetShape(xml, what)
  assertReferencesAreSane(xml, what)
  assertRowsAreUnique(xml, what)
}

/**
 * Parts other than these must survive an edit byte for byte. The workbook part
 * is here because a formula write marks it for recalculation; that it stays
 * untouched otherwise is pinned by a test in document.test.ts.
 */
const EXPECTED_TO_CHANGE = new Set([
  'xl/sharedStrings.xml',
  'xl/styles.xml',
  'xl/calcChain.xml',
  'xl/workbook.xml',
  '[Content_Types].xml',
])

export function assertOnlyTheSheetChanged(
  before: Uint8Array,
  after: Uint8Array,
  what: string,
): void {
  const original = readContainer(before)
  const written = readContainer(after)
  const sheets: string[] = []

  for (const [path, content] of original.parts) {
    if (EXPECTED_TO_CHANGE.has(path)) continue

    const other = written.parts.get(path)
    assert.ok(other !== undefined, `${what}: lost ${path}`)
    if (Buffer.compare(Buffer.from(content), Buffer.from(other)) === 0) continue

    assert.ok(
      path.toLowerCase().startsWith('xl/worksheets/'),
      `${what}: changed ${path}, which no edit should touch`,
    )
    sheets.push(path)
  }

  assert.ok(sheets.length <= 1, `${what}: changed several sheets: ${sheets.join(', ')}`)
}
