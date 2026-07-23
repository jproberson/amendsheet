import { collectRuns } from './collect.js'
import { printReport } from './report.js'

async function main() {
  const runs = await collectRuns(process.argv[2])
  for (const run of runs) printReport(run.title, run.results)
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
