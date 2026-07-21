// Opens our output in a spreadsheet application that is not this library.
//
// Every other check in this repo validates against our own reader, so a file we
// write wrongly and read back wrongly looks correct everywhere. This converts
// with LibreOffice instead, which is a real OOXML implementation with its own
// opinions about what is valid.
//
// The control matters: a fixture LibreOffice cannot open to begin with proves
// nothing about our writing. Each file is converted twice, before and after an
// edit, and only a file that converted before and fails after counts against us.
//
// Slow (a LibreOffice start per batch), so it is not part of verify.sh.
import { mkdir, mkdtemp, open, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { execFile } from 'node:child_process'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { createRequire } from 'node:module'

const run = promisify(execFile)
const require = createRequire(import.meta.url)
const { register } = require('tsx/esm/api')
const unregister = register()

const { readWorkbook } = await import('../src/lib/document.ts')

const SOFFICE = ['/Applications/LibreOffice.app/Contents/MacOS/soffice', 'soffice'].find(
  (path) => path === 'soffice' || existsSync(path),
)

const SENTINEL = 'amendsheet-external-check'
const FIXTURES = 'fixtures/real'

async function convert(files, into) {
  await mkdir(into, { recursive: true })
  if (files.length === 0) return
  // One start for the whole batch; LibreOffice takes seconds to boot.
  try {
    await run(SOFFICE, ['--headless', '--convert-to', 'csv', '--outdir', into, ...files], {
      timeout: 600_000,
      maxBuffer: 64 * 1024 * 1024,
    })
  } catch (error) {
    // A batch failure still leaves the files it managed; the caller checks what
    // actually landed rather than trusting the exit code.
    console.log(`  (converter exited non-zero: ${String(error).split('\n')[0]})`)
  }
}

const csvName = (file) => `${file.replace(/\.[^.]+$/, '')}.csv`

/**
 * The sentinel goes in the last row, so only the tail is read. One fixture
 * flattens to a twelve gigabyte csv all by itself, and reading that whole is
 * how this script first fell over.
 */
async function hasSentinel(path) {
  const { size } = await stat(path)
  const window = Math.min(size, 64 * 1024)
  const handle = await open(path, 'r')
  try {
    const buffer = Buffer.alloc(window)
    await handle.read(buffer, 0, window, size - window)
    return buffer.toString('utf8').includes(SENTINEL)
  } finally {
    await handle.close()
  }
}

const workDir = await mkdtemp(join(tmpdir(), 'amendsheet-external-'))
const originalDir = join(workDir, 'original')
const editedDir = join(workDir, 'edited')
const beforeOut = join(workDir, 'before')
const afterOut = join(workDir, 'after')

await mkdir(originalDir, { recursive: true })
await mkdir(editedDir, { recursive: true })

const files = (await readdir(FIXTURES)).filter((name) => name.endsWith('.xlsx'))
const edited = []
const refused = []

for (const file of files) {
  const bytes = new Uint8Array(await readFile(join(FIXTURES, file)))
  await writeFile(join(originalDir, file), bytes)

  try {
    const workbook = readWorkbook(bytes)
    const sheet = workbook.sheets[0]
    if (sheet === undefined) {
      refused.push(`${file}: no sheets`)
      continue
    }
    let lastRow = 0
    for (const cell of sheet.cells()) lastRow = Math.max(lastRow, cell.address.row)
    sheet.set(`A${lastRow + 1}`, SENTINEL)
    await writeFile(join(editedDir, file), workbook.toBytes())
    edited.push(file)
  } catch (error) {
    refused.push(`${file}: ${error instanceof Error ? error.message : String(error)}`)
  }
}

console.log(`converting ${files.length} originals`)
await convert(
  files.map((file) => join(originalDir, file)),
  beforeOut,
)
console.log(`converting ${edited.length} edited`)
await convert(
  edited.map((file) => join(editedDir, file)),
  afterOut,
)

const openedBefore = new Set(await readdir(beforeOut).catch(() => []))
const openedAfter = new Set(await readdir(afterOut).catch(() => []))

const broke = []
const sentinelMissing = []
let checked = 0

for (const file of edited) {
  const csv = csvName(file)
  if (!openedBefore.has(csv)) continue // no control, so it proves nothing
  checked++
  if (!openedAfter.has(csv)) {
    broke.push(file)
    continue
  }
  if (!(await hasSentinel(join(afterOut, csv)))) sentinelMissing.push(file)
}

console.log(`\n${'='.repeat(70)}`)
console.log('EXTERNAL READER: LibreOffice')
console.log('='.repeat(70))
console.log(`${files.length} fixtures, ${openedBefore.size} opened before editing`)
console.log(`${checked} had a working control and were checked after editing`)
console.log(`${broke.length} stopped opening after our edit`)
console.log(`${sentinelMissing.length} opened but lost the written value`)

for (const file of broke) console.log(`  BROKE: ${file}`)
for (const file of sentinelMissing) console.log(`  VALUE LOST: ${file}`)
for (const note of refused) console.log(`  refused: ${note}`)

await rm(workDir, { recursive: true, force: true })
await unregister()

if (broke.length > 0 || sentinelMissing.length > 0) process.exit(1)
