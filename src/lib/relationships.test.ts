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
  const relationships = readRelationships(RELS, 'rels')

  assert.equal(relationships.size, 3)
  assert.deepEqual(relationships.get('rId1'), {
    id: 'rId1',
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
    target: 'worksheets/sheet1.xml',
    external: false,
  })
})

test('marks external relationships', () => {
  const relationships = readRelationships(RELS, 'rels')

  assert.equal(relationships.get('rId3')?.external, true)
})

test('ignores elements that are not relationships', () => {
  const relationships = readRelationships('<Relationships><Other Id="x"/></Relationships>', 'rels')

  assert.equal(relationships.size, 0)
})

test('skips a relationship with no id, keeping the rest', () => {
  const relationships = readRelationships(
    '<Relationships><Relationship Target="a.xml"/><Relationship Id="rId1" Target="b.xml"/></Relationships>',
    'rels',
  )

  assert.equal(relationships.size, 1)
  assert.equal(relationships.get('rId1')?.target, 'b.xml')
})

test('rejects a relationship with no target', () => {
  assert.throws(
    () => readRelationships('<Relationships><Relationship Id="rId1"/></Relationships>', 'rels'),
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

      for (const relationship of readRelationships(
        new TextDecoder().decode(bytes),
        'rels',
      ).values()) {
        // External links and internal locations (a drawing pointing at
        // #Sheet1!A1) name no package part, so resolving them to one is not the
        // question being asked here.
        if (relationship.external || relationship.target.startsWith('#')) continue

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

  // A relationship may name a part the package never wrote — picture.xlsx points
  // rId3 at an xl/connections.xml it omits, and Excel opens it regardless — so a
  // reader cannot assume a target exists. What must hold is that resolution
  // produces a real path almost every time; a widespread miss is a resolver bug.
  assert.ok(
    dangling.length < checked / 100,
    `too many relationships resolved to a missing part: ${dangling.join(', ')}`,
  )
})
