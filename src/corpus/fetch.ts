import { mkdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

/**
 * Extends corpus/real with more files from the Apache POI test corpus.
 *
 * The corpus itself is committed, because a benchmark whose inputs can change
 * underneath it produces numbers nobody can compare. This script exists to grow
 * or refresh that set, not as a setup step.
 *
 * POI's files matter because no JavaScript library produced them: they were
 * collected from a decade of bug reports, written by real Excel versions,
 * LibreOffice, and third-party generators, and they carry the features a
 * synthetic corpus cannot fake — charts, pivot tables, embedded drawings,
 * threaded comments, themes.
 *
 * Pinned to a commit for the same reason the files are committed.
 */

const POI_COMMIT = '0d6d4872c491b1f230f51c6878e57407c60ae697'
const API = `https://api.github.com/repos/apache/poi/contents/test-data/spreadsheet?ref=${POI_COMMIT}`
const REAL_CORPUS_DIR = join(process.cwd(), 'corpus', 'real')
const MAX_BYTES = 600_000

/** Names hinting at document features that a round trip is likely to damage. */
const KEYWORDS = [
  'chart',
  'pivot',
  'image',
  'picture',
  'drawing',
  'comment',
  'conditional',
  'table',
  'style',
  'format',
  'date',
  'shared',
  'merge',
  'hyperlink',
  'validation',
  'theme',
  'filter',
  'freeze',
]

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

async function main() {
  const limit = Number(process.argv[2] ?? 60)

  const listing = await fetch(API, {
    headers: { accept: 'application/vnd.github+json', 'user-agent': 'xlsx-corpus-harness' },
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
        name.endsWith('.xlsx') &&
        entry.size < MAX_BYTES &&
        KEYWORDS.some((keyword) => name.includes(keyword))
      )
    })

  const selected = candidates.slice(0, limit)
  await mkdir(REAL_CORPUS_DIR, { recursive: true })

  let fetched = 0
  let skipped = 0
  for (const entry of selected) {
    try {
      const file = await fetch(entry.downloadUrl)
      if (!file.ok) throw new Error(`HTTP ${file.status}`)
      const bytes = new Uint8Array(await file.arrayBuffer())
      await writeFile(join(REAL_CORPUS_DIR, entry.name), bytes)
      fetched++
    } catch (error) {
      console.error(`  skipped ${entry.name}: ${describeError(error)}`)
      skipped++
    }
  }

  console.log(
    `Fetched ${fetched} of ${candidates.length} candidates${skipped ? ` (${skipped} skipped)` : ''}.`,
  )
  console.log('Commit any new files, and note the count in corpus/real/PROVENANCE.md.')
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
