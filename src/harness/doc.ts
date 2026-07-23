import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { collectRuns } from './collect.js'
import { buildComparison } from './markdown.js'

const OUTPUT = join(process.cwd(), 'COMPARISON.md')

async function main() {
  const runs = await collectRuns()
  await writeFile(OUTPUT, buildComparison(runs))
  console.log(`wrote ${OUTPUT}`)
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
