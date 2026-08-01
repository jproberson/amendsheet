import { readdir, readFile } from 'node:fs/promises'
import { join, relative } from 'node:path'
import { measureRoundTrip } from './compare.js'
import { amendsheetAdapter } from '../adapters/amendsheet.js'
import { exceljsAdapter } from '../adapters/exceljs.js'
import type { Adapter, RoundTripResult } from './types.js'

const FIXTURES_DIR = join(process.cwd(), 'fixtures')

/** Register additional libraries here to benchmark them side by side. */
const ADAPTERS: Adapter[] = [amendsheetAdapter, exceljsAdapter]

export interface Run {
  title: string
  results: RoundTripResult[]
}

async function findFiles(dir: string): Promise<string[]> {
  const found: string[] = []
  const entries = await readdir(dir, { withFileTypes: true }).catch(() => [])
  for (const entry of entries) {
    const full = join(dir, entry.name)
    if (entry.isDirectory()) found.push(...(await findFiles(full)))
    else if (/\.xlsm?x?$/.test(entry.name) && !entry.name.startsWith('~$')) found.push(full)
  }
  return found.sort()
}

async function measureAll(
  adapter: Adapter,
  files: string[],
  edited: boolean,
): Promise<RoundTripResult[]> {
  const results: RoundTripResult[] = []
  for (const file of files) {
    const bytes = new Uint8Array(await readFile(file))
    results.push(await measureRoundTrip(adapter, relative(FIXTURES_DIR, file), bytes, edited))
  }
  return results
}

/**
 * Run every adapter over the fixtures and return one entry per measured pass:
 * a round trip, and — for a library that edits — the same files after one cell
 * changes. Shared by the console report and the Markdown comparison so they can
 * never disagree.
 */
export async function collectRuns(only?: string): Promise<Run[]> {
  const all = await findFiles(FIXTURES_DIR)
  if (all.length === 0) {
    throw new Error(`No .xlsx files under ${FIXTURES_DIR}. Run: npm run fixtures`)
  }
  const files = only ? all.filter((file) => file.includes(only)) : all

  const runs: Run[] = []
  for (const adapter of ADAPTERS) {
    runs.push({ title: adapter.name, results: await measureAll(adapter, files, false) })
    if (adapter.edit !== undefined) {
      runs.push({
        title: `${adapter.name}, after editing`,
        results: await measureAll(adapter, files, true),
      })
    }
  }
  return runs
}
