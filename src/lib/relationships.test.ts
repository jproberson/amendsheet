import assert from 'node:assert/strict'
import { test } from 'node:test'
import { readRelationships, resolveTarget } from './relationships.js'

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`

test('reads relationships by id', () => {
  const relationships = readRelationships(RELS)

  assert.equal(relationships.size, 3)
  assert.deepEqual(relationships.get('rId1'), {
    id: 'rId1',
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
    target: 'worksheets/sheet1.xml',
    external: false,
  })
})

test('marks external relationships', () => {
  const relationships = readRelationships(RELS)

  assert.equal(relationships.get('rId3')?.external, true)
})

test('ignores elements that are not relationships', () => {
  const relationships = readRelationships('<Relationships><Other Id="x"/></Relationships>')

  assert.equal(relationships.size, 0)
})

test('rejects a relationship with no id', () => {
  assert.throws(
    () => readRelationships('<Relationships><Relationship Target="a.xml"/></Relationships>'),
    /Relationship is missing Id/,
  )
})

test('rejects a relationship with no target', () => {
  assert.throws(
    () => readRelationships('<Relationships><Relationship Id="rId1"/></Relationships>'),
    /rId1 is missing Target/,
  )
})

test('resolves a target against the folder of the part that owns it', () => {
  assert.equal(
    resolveTarget('xl/workbook.xml', 'worksheets/sheet1.xml'),
    'xl/worksheets/sheet1.xml',
  )
})

test('resolves a target from the package root', () => {
  assert.equal(resolveTarget('', 'xl/workbook.xml'), 'xl/workbook.xml')
})

test('resolves a target that walks up a folder', () => {
  assert.equal(
    resolveTarget('xl/worksheets/sheet1.xml', '../drawings/drawing1.xml'),
    'xl/drawings/drawing1.xml',
  )
})

test('resolves an absolute target', () => {
  assert.equal(resolveTarget('xl/worksheets/sheet1.xml', '/xl/styles.xml'), 'xl/styles.xml')
})

test('ignores a redundant current folder segment', () => {
  assert.equal(resolveTarget('xl/workbook.xml', './styles.xml'), 'xl/styles.xml')
})

test('resolves fixtures relationships to package paths, dangling ones included', async () => {
  const { readdir, readFile } = await import('node:fs/promises')
  const { readContainer } = await import('./container.js')

  const files = (await readdir('fixtures/real')).filter((name) => name.endsWith('.xlsx'))
  const dangling: string[] = []
  let checked = 0

  for (const file of files) {
    const container = readContainer(new Uint8Array(await readFile(`fixtures/real/${file}`)))

    for (const [path, bytes] of container.parts) {
      if (!path.endsWith('.rels')) continue

      // xl/_rels/workbook.xml.rels describes xl/workbook.xml
      const owner = path.replace('_rels/', '').replace(/\.rels$/, '')

      for (const relationship of readRelationships(new TextDecoder().decode(bytes)).values()) {
        if (relationship.external) continue

        const resolved = resolveTarget(owner, relationship.target)
        assert.ok(!resolved.startsWith('/'), `${file}: ${resolved} kept a leading slash`)
        assert.ok(!resolved.includes('..'), `${file}: ${resolved} kept a parent segment`)
        assert.notEqual(resolved, '', `${file}: ${relationship.target} resolved to nothing`)

        if (!container.parts.has(resolved)) dangling.push(`${file} -> ${resolved}`)
        checked++
      }
    }
  }

  assert.ok(checked > 100, `expected a meaningful number of relationships, got ${checked}`)

  // picture.xlsx declares xl/connections.xml in [Content_Types].xml and points
  // rId3 at it, but never writes the part. Excel opens the file regardless, so
  // reading one cannot assume a relationship target exists.
  assert.deepEqual(dangling, ['picture.xlsx -> xl/connections.xml'])
})
