import type { RoundTripResult } from './types.js'

const fit = (text: string, width: number) =>
  text.length > width ? `${text.slice(0, width - 1)}…` : text.padEnd(width)

export function printReport(adapterName: string, results: RoundTripResult[]): void {
  console.log(`\n${'='.repeat(78)}`)
  console.log(`ROUND-TRIP FIDELITY: ${adapterName}`)
  console.log(`Property tested: read(f) -> write() must preserve every part and value`)
  console.log('='.repeat(78))

  console.log(`\n${fit('FILE', 34)}${fit('VERDICT', 10)}${fit('PARTS LOST', 12)}CELLS`)
  console.log('-'.repeat(78))

  for (const r of results) {
    const verdict = r.error ? 'ERROR' : r.ok ? 'ok' : 'DATA LOSS'
    const cells = r.error ? '-' : `${r.cellsTotal - r.cellsLost - r.cellsChanged}/${r.cellsTotal}`
    console.log(fit(r.file, 34) + fit(verdict, 10) + fit(String(r.partsLost.length), 12) + cells)
  }

  const failures = results.filter((r) => !r.ok)
  if (failures.length > 0) {
    console.log(`\n${'-'.repeat(78)}\nDETAIL\n${'-'.repeat(78)}`)
    for (const r of failures) {
      console.log(`\n${r.file}`)
      if (r.error) {
        console.log(`  ERROR: ${r.error}`)
        continue
      }
      if (r.partsLost.length) {
        console.log(`  parts dropped entirely:`)
        for (const p of r.partsLost) console.log(`    - ${p}`)
      }
      if (r.featureLoss.length) {
        console.log(`  features degraded:`)
        for (const f of r.featureLoss) {
          console.log(`    - ${f.feature}: ${f.before} -> ${f.after}`)
        }
      }
      if (r.cellsLost) console.log(`  cell values lost: ${r.cellsLost}`)
      if (r.cellsChanged) console.log(`  cell values altered: ${r.cellsChanged}`)
      if (r.partsAdded.length) {
        console.log(`  parts added: ${r.partsAdded.join(', ')}`)
      }
    }
  }

  const ok = results.filter((r) => r.ok).length
  const errored = results.filter((r) => r.error).length
  const lossy = results.length - ok - errored

  console.log(`\n${'='.repeat(78)}`)
  console.log(
    `SUMMARY  ${ok} clean  |  ${lossy} lossy  |  ${errored} failed to process  (of ${results.length})`,
  )

  const featureTotals = new Map<string, number>()
  for (const r of results) {
    for (const f of r.featureLoss) {
      featureTotals.set(f.feature, (featureTotals.get(f.feature) ?? 0) + 1)
    }
  }
  if (featureTotals.size > 0) {
    console.log(`\nMost frequently damaged features:`)
    const ranked = [...featureTotals].sort((a, b) => b[1] - a[1])
    for (const [feature, count] of ranked) {
      console.log(`  ${fit(feature, 26)} damaged in ${count} file(s)`)
    }
  }
  console.log('='.repeat(78))
}
