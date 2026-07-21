// Reads and edits a real file using the built output, the way a consumer would.
// Run against every Node version the package claims to support.
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import { createRequire } from 'node:module'

import { readWorkbook } from '../dist/index.js'

const bytes = new Uint8Array(await readFile('fixtures/real/WithChart.xlsx'))

const workbook = readWorkbook(bytes)
assert.ok(workbook.sheets.length > 0, 'no sheets were read')

const cells = [...workbook.sheets[0].cells()]
assert.ok(cells.length > 0, 'no cells were read')

workbook.sheets[0].set('A1', 'smoke test')
const saved = workbook.toBytes()

const reopened = readWorkbook(saved)
const edited = [...reopened.sheets[0].cells()].find((cell) => cell.reference === 'A1')
assert.deepEqual(edited.value, { kind: 'text', value: 'smoke test' })

const require = createRequire(import.meta.url)
const commonjs = require('../dist/index.cjs')
assert.equal(typeof commonjs.readWorkbook, 'function', 'require() did not work')

console.log(`ok on node ${process.version}: read ${cells.length} cells, edited, and required`)
