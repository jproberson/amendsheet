import { readdir, readFile } from 'node:fs/promises'
import { join, relative } from 'node:path'
import { measureRoundTrip } from './compare.js'
import { printReport } from './report.js'
import { exceljsAdapter } from '../adapters/exceljs.js'
import { amendsheetAdapter } from '../adapters/amendsheet.js'
import type { Adapter, RoundTripResult } from './types.js'

const FIXTURES_DIR = join(process.cwd(), 'fixtures')

/** Register additional libraries here to benchmark them side by side. */
const ADAPTERS: Adapter[] = [amendsheetAdapter, exceljsAdapter]

async function findFiles(dir: string): Promise<string[]> {
  const found: string[] = []
  const entries = await readdir(dir, { withFileTypes: true }).catch(() => [])
  for (const entry of entries) {
    const full = join(dir, entry.name)
    if (entry.isDirectory()) found.push(...(await findFiles(full)))
    else if (entry.name.endsWith('.xlsx') && !entry.name.startsWith('~$')) found.push(full)
  }
  return found.sort()
}

async function main() {
  const files = await findFiles(FIXTURES_DIR)
  if (files.length === 0) {
    console.error(`No .xlsx files under ${FIXTURES_DIR}. Run: npm run fixtures`)
    process.exit(1)
  }

  const only = process.argv[2]
  const selected = only ? files.filter((f) => f.includes(only)) : files

  for (const adapter of ADAPTERS) {
    const results: RoundTripResult[] = []
    for (const file of selected) {
      const bytes = new Uint8Array(await readFile(file))
      results.push(await measureRoundTrip(adapter, relative(FIXTURES_DIR, file), bytes))
    }
    printReport(adapter.name, results)

    if (adapter.edit === undefined) continue

    const edited: RoundTripResult[] = []
    for (const file of files) {
      const bytes = new Uint8Array(await readFile(file))
      edited.push(await measureRoundTrip(adapter, relative(FIXTURES_DIR, file), bytes, true))
    }
    printReport(`${adapter.name}, after editing one cell`, edited)
  }
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
