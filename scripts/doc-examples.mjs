// Typechecks the ```ts examples in README.md against the real exports, so a
// renamed or removed part of the API breaks the docs rather than drifting from
// them silently. It is the docs' equivalent of the browser smoke test: the
// claim is proven by compilation, not by a grep.
//
// The examples are illustrative, so a small prelude supplies the values they
// assume — `bytes`, `value`, a `writeFile` — and every example may use them.
// Two conventions keep the prelude fixed rather than per-example:
//   - a usage example declares its own workbook and sheet, and imports from
//     'amendsheet' or leans on the prelude's import;
//   - a type example (one that begins `type` or `interface`) is checked for
//     assignability both ways against the export of the same name, so a copied
//     type that falls out of step with the real one fails.

import { spawnSync } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = fileURLToPath(new URL('..', import.meta.url))

const usagePrelude = [
  "import { createWorkbook, readWorkbook } from 'amendsheet'",
  "import type { CellValue } from 'amendsheet'",
  'declare const bytes: Uint8Array',
  'declare const value: CellValue',
  'declare function writeFile(path: string, data: Uint8Array): Promise<void>',
  '',
  '',
].join('\n')

const isTypeExample = (body) => /^\s*(export\s+)?(type|interface)\b/.test(body)

function wrapType(body) {
  const match = body.match(/(?:export\s+)?(?:type|interface)\s+([A-Za-z0-9_$]+)/)
  if (match === null) return undefined
  const name = match[1]
  return [
    `import type { ${name} as Real } from 'amendsheet'`,
    body,
    `const forward = (r: Real): ${name} => r`,
    `const back = (d: ${name}): Real => d`,
    'export {}',
    '',
  ].join('\n')
}

function wrapUsage(body) {
  const stripped = body.replace(/^\s*import\b[^\n]*from\s*['"]amendsheet['"];?[ \t]*$/gm, '')
  return `${usagePrelude}${stripped}\nexport {}\n`
}

function examplesIn(markdown) {
  const fences = [...markdown.matchAll(/```ts\n([\s\S]*?)```/g)]
  return fences.map((fence) => fence[1])
}

async function main() {
  const readme = await readFile(join(root, 'README.md'), 'utf8')
  const examples = examplesIn(readme)
  if (examples.length === 0) {
    console.log('SKIPPED: no ts examples in README.md')
    return
  }

  const dir = await mkdtemp(join(tmpdir(), 'doc-examples-'))
  try {
    const files = []
    for (const [index, body] of examples.entries()) {
      const wrapped = isTypeExample(body) ? wrapType(body) : wrapUsage(body)
      if (wrapped === undefined) {
        console.log(`FAIL: README example ${index + 1} declares no named type to check`)
        process.exitCode = 1
        return
      }
      const name = `example-${index + 1}.ts`
      await writeFile(join(dir, name), wrapped)
      files.push(name)
    }

    const tsconfig = {
      extends: join(root, 'tsconfig.json'),
      compilerOptions: {
        noUncheckedIndexedAccess: false,
        types: [],
        lib: ['ES2022', 'DOM'],
        baseUrl: '.',
        paths: { amendsheet: [join(root, 'src/index')] },
      },
      include: ['*.ts'],
    }
    await writeFile(join(dir, 'tsconfig.json'), JSON.stringify(tsconfig))

    const result = spawnSync('npx', ['tsc', '-p', dir], { encoding: 'utf8' })
    if (result.status === 0) {
      console.log(`PASSED: ${files.length} README examples typecheck against the exports`)
      return
    }

    console.log('FAIL: a README example does not match the exports')
    process.stdout.write(result.stdout ?? '')
    process.stdout.write(result.stderr ?? '')
    process.exitCode = 1
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

await main()
