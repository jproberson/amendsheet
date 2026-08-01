import { collectRuns } from './collect.js'
import { printReport } from './report.js'
import { amendsheetAdapter } from '../adapters/amendsheet.js'

async function main() {
  const runs = await collectRuns(process.argv[2])
  for (const run of runs) printReport(run.title, run.results)

  // This library's own passes decide the exit code, so a regression fails the
  // command's status and not only a line in a report a reader might skim past.
  // A rewritten part counts too: on the round-trip pass it breaks the promise to
  // leave untouched parts alone. The comparison adapters are expected to lose
  // data, so they never gate.
  const mine = runs.filter((run) => run.title.startsWith(amendsheetAdapter.name))
  const regressions = mine.flatMap((run) =>
    run.results.filter((result) => !result.ok || result.partsChanged.length > 0),
  )
  if (regressions.length > 0) {
    console.error(`\nharness: ${regressions.length} amendsheet file(s) not clean; see DETAIL above`)
    process.exitCode = 1
  }
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
