// Edits every fixture and checks the output against the integrity rules Excel
// repairs a file for: a relationship whose target is missing, a part with no
// content type, and a table whose shape is inconsistent. It is not Excel, but
// it is the mechanical core of what Excel checks, run over every fixture rather
// than a handful opened by hand.

import { readdir, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { readContainer } from '../src/lib/container.ts'
import { readWorkbook } from '../src/lib/document.ts'
import { readRelationships, resolveTarget } from '../src/lib/relationships.ts'
import { readXml } from '../src/lib/xml.ts'

const FIXTURES = 'fixtures'

async function findFixtures(dir) {
  const found = []
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name)
    if (entry.isDirectory()) found.push(...(await findFixtures(full)))
    else if (/\.xls[xm]$/.test(entry.name)) found.push(full)
  }
  return found
}

const decode = (bytes) => new TextDecoder('utf-8', { fatal: false }).decode(bytes)

/** Every part name a content type covers, by Override name and by Default extension. */
function contentTypeCoverage(xml) {
  const defaults = new Set()
  const overrides = new Set()
  for (const event of readXml(xml)) {
    if (event.kind !== 'open') continue
    if (event.localName === 'Default') {
      const ext = event.attributes.get('Extension')
      if (ext !== undefined) defaults.add(ext.toLowerCase())
    }
    if (event.localName === 'Override') {
      const name = event.attributes.get('PartName')
      if (name !== undefined) overrides.add(name)
    }
  }
  return { defaults, overrides }
}

function checkContentTypes(parts, problems) {
  const typesXml = parts.get('[Content_Types].xml')
  if (typesXml === undefined) {
    problems.push('no [Content_Types].xml')
    return
  }
  const { defaults, overrides } = contentTypeCoverage(decode(typesXml))
  for (const path of parts.keys()) {
    if (path === '[Content_Types].xml') continue
    const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
    if (overrides.has(`/${path}`) || defaults.has(extension)) continue
    problems.push(`part has no content type: ${path}`)
  }
}

function checkRelationships(parts, problems) {
  for (const [path, bytes] of parts) {
    if (!path.endsWith('.rels')) continue
    const owner = path.replace('_rels/', '').replace(/\.rels$/, '')
    let relationships
    try {
      relationships = readRelationships(decode(bytes), path)
    } catch (error) {
      problems.push(`unreadable relationships in ${path}: ${error.message}`)
      continue
    }
    for (const relationship of relationships.values()) {
      if (relationship.external) continue
      const target = resolveTarget(owner, relationship.target)
      if (!parts.has(target)) {
        problems.push(`dangling relationship ${relationship.id} in ${path} -> ${target}`)
      }
    }
  }
}

function checkTables(parts, problems) {
  for (const [path, bytes] of parts) {
    if (!/^xl\/tables\/[^/]+\.xml$/.test(path)) continue
    const xml = decode(bytes)
    let declaredColumns
    let seenColumns = 0
    const ids = new Set()
    const names = new Set()
    for (const event of readXml(xml)) {
      if (event.kind !== 'open') continue
      if (event.localName === 'tableColumns') {
        const count = Number(event.attributes.get('count'))
        if (Number.isInteger(count)) declaredColumns = count
      }
      if (event.localName === 'tableColumn') {
        seenColumns++
        const id = event.attributes.get('id')
        const name = event.attributes.get('name')?.toLowerCase()
        if (id !== undefined && ids.has(id)) problems.push(`${path}: duplicate column id ${id}`)
        if (name !== undefined && names.has(name)) problems.push(`${path}: duplicate column name ${name}`)
        if (id !== undefined) ids.add(id)
        if (name !== undefined) names.add(name)
      }
    }
    if (declaredColumns !== undefined && declaredColumns !== seenColumns) {
      problems.push(`${path}: tableColumns count ${declaredColumns} but ${seenColumns} columns`)
    }
  }
}

function validate(bytes) {
  const problems = []
  let parts
  try {
    parts = readContainer(bytes).parts
  } catch (error) {
    return [`not a readable package: ${error.message}`]
  }
  checkContentTypes(parts, problems)
  checkRelationships(parts, problems)
  checkTables(parts, problems)
  return problems
}

function editValue(workbook) {
  const sheet = workbook.sheets[0]
  if (sheet === undefined) return
  let lastRow = 0
  for (const cell of sheet.cells()) lastRow = Math.max(lastRow, cell.address.row)
  sheet.set(`A${lastRow + 1}`, 'validate-opc')
}

function editFormula(workbook) {
  const sheet = workbook.sheets[0]
  if (sheet === undefined) return
  let lastRow = 0
  for (const cell of sheet.cells()) lastRow = Math.max(lastRow, cell.address.row)
  sheet.set(`A${lastRow + 1}`, { formula: 'B1+1' })
}

async function main() {
  const files = (await findFixtures(FIXTURES)).sort()
  let checked = 0
  let failed = 0

  for (const file of files) {
    const bytes = new Uint8Array(await readFile(file))
    // Some source fixtures ship with their own quirks (picture.xlsx points a
    // relationship at a part it never writes, and Excel opens it anyway). Only
    // a problem our edit introduces, not one already in the file, is a failure.
    const baseline = new Set(validate(bytes))

    for (const [label, edit] of [
      ['value below', editValue],
      ['formula', editFormula],
    ]) {
      let out
      try {
        const workbook = readWorkbook(bytes)
        edit(workbook)
        out = workbook.toBytes()
      } catch (error) {
        console.log(`FAIL ${file} [${label}] threw: ${error.message}`)
        failed++
        continue
      }
      const introduced = validate(out).filter((problem) => !baseline.has(problem))
      checked++
      if (introduced.length > 0) {
        failed++
        console.log(`FAIL ${file} [${label}]`)
        for (const problem of introduced) console.log(`     ${problem}`)
      }
    }
  }

  console.log(`\n${checked - failed}/${checked} package checks clean across ${files.length} fixtures`)
  if (failed > 0) process.exitCode = 1
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
