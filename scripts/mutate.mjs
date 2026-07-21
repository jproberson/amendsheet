// Deliberately breaks the library one edit at a time and checks that some test
// notices. Coverage says a line ran; this says the line mattered.
//
// Slow by nature, so it is not part of verify.sh. Run it when the test suite
// has changed shape, or before a release.
import { execFileSync } from 'node:child_process'
import { readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

const LIB = 'src/lib'

/** Each rule is a search and a replacement that should break behaviour. */
const RULES = [
  { name: 'flip >= to >', find: / >= /g, put: ' > ' },
  { name: 'flip <= to <', find: / <= /g, put: ' < ' },
  { name: 'flip === to !==', find: / === /g, put: ' !== ' },
  { name: 'flip && to ||', find: / && /g, put: ' || ' },
  { name: 'off by one on +1', find: /\+ 1\b/g, put: '+ 2' },
  { name: 'drop a continue', find: /^(\s+)continue$/gm, put: '$1;' },
  { name: 'return early', find: /^(\s+)if \((.+)\) return (.+)$/gm, put: '$1if (!($2)) return $3' },
]

const files = readdirSync(LIB)
  .filter((name) => name.endsWith('.ts') && !name.endsWith('.test.ts'))
  .map((name) => join(LIB, name))

function testsPass() {
  try {
    execFileSync('npm', ['test'], { stdio: 'ignore' })
    return true
  } catch {
    return false
  }
}

const limit = Number(process.argv[2] ?? 40)
const survivors = []
let applied = 0

console.log(`Mutating ${files.length} files, up to ${limit} mutations.\n`)

outer: for (const file of files) {
  const original = readFileSync(file, 'utf8')

  for (const rule of RULES) {
    const matches = [...original.matchAll(rule.find)]
    if (matches.length === 0) continue

    // One mutation per rule per file, at the first site, to keep the run finite.
    const at = matches[0].index
    const mutated =
      original.slice(0, at) +
      matches[0][0].replace(rule.find, rule.put) +
      original.slice(at + matches[0][0].length)
    if (mutated === original) continue

    writeFileSync(file, mutated)
    const survived = testsPass()
    writeFileSync(file, original)

    applied++
    const line = original.slice(0, at).split('\n').length
    if (survived) {
      survivors.push(`${file}:${line} ${rule.name}`)
      console.log(`SURVIVED  ${file}:${line}  ${rule.name}`)
    } else {
      console.log(`caught    ${file}:${line}  ${rule.name}`)
    }

    if (applied >= limit) break outer
  }
}

console.log(`\n${applied - survivors.length} of ${applied} mutations were caught.`)
if (survivors.length > 0) {
  console.log('\nSurvived, so read each one and decide which it is:')
  console.log('  a gap, where the behaviour is real but nothing asserts it, or')
  console.log('  an equivalent mutation, which changes nothing and no test can catch.')
  for (const survivor of survivors) console.log(`  ${survivor}`)
}
process.exit(survivors.length > 0 ? 1 : 0)
