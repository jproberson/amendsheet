import { mkdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { measureRoundTrip } from '../harness/compare.js'
import { amendsheetAdapter } from '../adapters/amendsheet.js'

/**
 * Extends fixtures/real with more files from the Apache POI test data.
 *
 * The corpus is committed, because a benchmark whose inputs can change underneath
 * it produces numbers nobody can compare. This script grows or refreshes that
 * set, it is not a setup step.
 *
 * POI's files matter because no JavaScript library produced them: collected from
 * a decade of bug reports, written by real Excel versions, LibreOffice, and
 * third-party generators, they carry the features and structural oddities a
 * synthetic fixture cannot fake.
 *
 * Selection is breadth-first: every workbook under the size cap, minus files
 * named as a deliberately-broken input (a fuzzer minimisation, a decompression
 * bomb, a corrupt package). Those belong in an error-path test that asserts a
 * located failure, not in a fidelity corpus, and processing one here could hang.
 *
 * Each candidate is then round-tripped and edited through the same measurement
 * the harness runs, and kept only if both passes are clean — so a file that
 * would turn the harness red can never land in the corpus by accident. A file
 * that reveals a real gap (Strict OOXML today) stays out until the gap is fixed,
 * at which point it round-trips clean and is picked up on the next run.
 *
 * Pinned to a commit for the same reason the files are committed.
 */

const POI_COMMIT = '0d6d4872c491b1f230f51c6878e57407c60ae697'
const API = `https://api.github.com/repos/apache/poi/contents/test-data/spreadsheet?ref=${POI_COMMIT}`
const REAL_FIXTURES_DIR = join(process.cwd(), 'fixtures', 'real')
const MAX_BYTES = 600_000

/** Names of deliberately-broken inputs: a clean round trip is not the contract
 * for these, and a decompression bomb must not be fed to the parser at all. */
const BROKEN = /corrupt|malformed|clusterfuzz|xmlbomb|crash-|fuzz|poc-|invalid|truncat/i

interface DirectoryEntry {
  name: string
  size: number
  downloadUrl: string
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

/** The GitHub contents API is untrusted input, so shape is checked rather than assumed. */
function toDirectoryEntry(value: unknown): DirectoryEntry | null {
  if (!isObject(value)) return null
  const { name, size, download_url: downloadUrl } = value
  if (typeof name !== 'string' || typeof size !== 'number') return null
  if (typeof downloadUrl !== 'string') return null
  return { name, size, downloadUrl }
}

function describeError(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/** Clean means both harness passes preserved every part, feature and value, and
 * rewrote nothing they were meant to leave alone. Returns why, when not. */
async function reasonUnclean(name: string, bytes: Uint8Array): Promise<string | undefined> {
  for (const edited of [false, true]) {
    const result = await measureRoundTrip(amendsheetAdapter, name, bytes, edited)
    const pass = edited ? 'edit' : 'round-trip'
    if (result.error) return result.error
    if (result.partsLost.length) return `${pass}: lost ${result.partsLost.join(', ')}`
    if (result.featureLoss.length) {
      return `${pass}: ${result.featureLoss.map((f) => `${f.feature} ${f.before}->${f.after}`).join(', ')}`
    }
    if (result.cellsLost || result.cellsChanged) {
      return `${pass}: ${result.cellsLost} cells lost, ${result.cellsChanged} changed`
    }
    if (result.partsChanged.length) return `${pass}: rewrote ${result.partsChanged.join(', ')}`
  }
  return undefined
}

async function main() {
  const limit = Number(process.argv[2] ?? Number.POSITIVE_INFINITY)

  const listing = await fetch(API, {
    headers: { accept: 'application/vnd.github+json', 'user-agent': 'xlsx-fixtures-harness' },
  })
  if (!listing.ok) {
    console.error(`GitHub API returned ${listing.status}. Rate limited? Try again shortly.`)
    process.exit(1)
  }

  const payload: unknown = await listing.json()
  if (!Array.isArray(payload)) {
    console.error('GitHub API returned an unexpected shape; expected an array of entries.')
    process.exit(1)
  }

  const candidates = payload
    .map(toDirectoryEntry)
    .filter((entry): entry is DirectoryEntry => entry !== null)
    .filter((entry) => {
      const name = entry.name.toLowerCase()
      return (
        (name.endsWith('.xlsx') || name.endsWith('.xlsm')) &&
        entry.size < MAX_BYTES &&
        !BROKEN.test(name)
      )
    })
    .slice(0, limit)

  await mkdir(REAL_FIXTURES_DIR, { recursive: true })

  let kept = 0
  const skipped: string[] = []
  for (const entry of candidates) {
    let bytes: Uint8Array
    try {
      const file = await fetch(entry.downloadUrl)
      if (!file.ok) throw new Error(`HTTP ${file.status}`)
      bytes = new Uint8Array(await file.arrayBuffer())
    } catch (error) {
      skipped.push(`${entry.name}: download failed (${describeError(error)})`)
      continue
    }

    const unclean = await reasonUnclean(entry.name, bytes)
    if (unclean !== undefined) {
      skipped.push(`${entry.name}: ${unclean}`)
      continue
    }

    await writeFile(join(REAL_FIXTURES_DIR, entry.name), bytes)
    kept++
  }

  console.log(`Kept ${kept} of ${candidates.length} candidates.`)
  if (skipped.length > 0) {
    console.log(`\nSkipped ${skipped.length} that would not round-trip clean:`)
    for (const line of skipped) console.log(`  ${line}`)
  }
  console.log('\nCommit any new files, and note the count in fixtures/real/PROVENANCE.md.')
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
