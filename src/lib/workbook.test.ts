import assert from 'node:assert/strict'
import { readFile, readdir } from 'node:fs/promises'
import { test } from 'node:test'
import { writeContainer } from './container.js'
import { openWorkbook } from './workbook.js'

const encode = (text: string) => new TextEncoder().encode(text)

const ROOT_RELS = `<Relationships>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK_RELS = `<Relationships>
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`

function packageOf(workbookXml: string, overrides: Record<string, string> = {}): Uint8Array {
  const parts: Record<string, Uint8Array> = {
    '_rels/.rels': encode(ROOT_RELS),
    'xl/workbook.xml': encode(workbookXml),
    'xl/_rels/workbook.xml.rels': encode(WORKBOOK_RELS),
    'xl/worksheets/sheet1.xml': encode('<worksheet/>'),
    'xl/worksheets/sheet2.xml': encode('<worksheet/>'),
  }
  for (const [path, content] of Object.entries(overrides)) parts[path] = encode(content)
  return writeContainer({ parts: new Map(Object.entries(parts)) })
}

const TWO_SHEETS = `<workbook>
<sheets>
<sheet name="Data" sheetId="1" r:id="rId1"/>
<sheet name="Notes" sheetId="2" r:id="rId2" state="hidden"/>
</sheets>
</workbook>`

test('lists sheets in document order', () => {
  const workbook = openWorkbook(packageOf(TWO_SHEETS))

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Data', 'Notes'],
  )
})

test('resolves each sheet to its part path', () => {
  const workbook = openWorkbook(packageOf(TWO_SHEETS))

  assert.equal(workbook.sheets[0]?.path, 'xl/worksheets/sheet1.xml')
  assert.equal(workbook.sheets[1]?.path, 'xl/worksheets/sheet2.xml')
})

test('reads sheet visibility', () => {
  const workbook = openWorkbook(packageOf(TWO_SHEETS))

  assert.equal(workbook.sheets[0]?.state, 'visible')
  assert.equal(workbook.sheets[1]?.state, 'hidden')
})

test('defaults to the 1900 date system', () => {
  assert.equal(openWorkbook(packageOf(TWO_SHEETS)).date1904, false)
})

test('reads the 1904 date system flag', () => {
  const xml = `<workbook><workbookPr date1904="1"/><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>`

  assert.equal(openWorkbook(packageOf(xml)).date1904, true)
})

test('accepts the 1904 flag written as true', () => {
  const xml = `<workbook><workbookPr date1904="true"/><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>`

  assert.equal(openWorkbook(packageOf(xml)).date1904, true)
})

test('keeps the container so unread parts survive', () => {
  const workbook = openWorkbook(packageOf(TWO_SHEETS))

  assert.equal(workbook.container.parts.has('xl/worksheets/sheet2.xml'), true)
})

test('rejects a package with no root relationships', () => {
  const bytes = writeContainer({ parts: new Map([['xl/workbook.xml', encode('<workbook/>')]]) })

  assert.throws(() => openWorkbook(bytes), /_rels\/\.rels/)
})

test('rejects a package with no workbook relationship', () => {
  const bytes = writeContainer({
    parts: new Map([['_rels/.rels', encode('<Relationships/>')]]),
  })

  assert.throws(() => openWorkbook(bytes), /no workbook/i)
})

test('rejects a workbook part that is missing', () => {
  const bytes = writeContainer({ parts: new Map([['_rels/.rels', encode(ROOT_RELS)]]) })

  assert.throws(() => openWorkbook(bytes), /xl\/workbook\.xml/)
})

test('skips a sheet whose relationship does not resolve', () => {
  const xml = `<workbook><sheets><sheet name="Ghost" sheetId="1" r:id="rId99"/><sheet name="Real" sheetId="2" r:id="rId1"/></sheets></workbook>`

  const workbook = openWorkbook(packageOf(xml))

  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Real'],
  )
})

test('opens every workbook in the corpus', async () => {
  const files = (await readdir('corpus/real')).filter((name) => name.endsWith('.xlsx'))
  const failures: string[] = []
  let sheets = 0

  for (const file of files) {
    try {
      const workbook = openWorkbook(new Uint8Array(await readFile(`corpus/real/${file}`)))
      assert.ok(workbook.sheets.length > 0, `${file} reported no sheets`)
      for (const sheet of workbook.sheets) {
        assert.ok(
          workbook.container.parts.has(sheet.path),
          `${file}: ${sheet.name} points at missing ${sheet.path}`,
        )
        sheets++
      }
    } catch (error) {
      failures.push(`${file}: ${error instanceof Error ? error.message : String(error)}`)
    }
  }

  assert.deepEqual(failures, [])
  assert.ok(sheets > 60, `expected many sheets across the corpus, got ${sheets}`)
})

test('reads a very hidden sheet', () => {
  const xml = `<workbook><sheets><sheet name="A" sheetId="1" r:id="rId1" state="veryHidden"/></sheets></workbook>`

  assert.equal(openWorkbook(packageOf(xml)).sheets[0]?.state, 'veryHidden')
})

test('accepts a relationship prefix other than r', () => {
  const xml = `<workbook><sheets><sheet name="A" sheetId="1" rel:id="rId1"/></sheets></workbook>`

  assert.equal(openWorkbook(packageOf(xml)).sheets[0]?.path, 'xl/worksheets/sheet1.xml')
})

test('skips a sheet with no relationship attribute at all', () => {
  const xml = `<workbook><sheets><sheet name="A" sheetId="1"/></sheets></workbook>`

  assert.deepEqual(openWorkbook(packageOf(xml)).sheets, [])
})

test('skips a sheet whose relationship is external', () => {
  const rels = `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="http://example.com/s.xml" TargetMode="External"/></Relationships>`
  const xml = `<workbook><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>`

  const workbook = openWorkbook(packageOf(xml, { 'xl/_rels/workbook.xml.rels': rels }))

  assert.deepEqual(workbook.sheets, [])
})

test('defaults name and sheetId when the attributes are absent', () => {
  const xml = `<workbook><sheets><sheet r:id="rId1"/></sheets></workbook>`

  const sheet = openWorkbook(packageOf(xml)).sheets[0]

  assert.equal(sheet?.name, '')
  assert.equal(sheet?.sheetId, '')
})

test('reads a workbook that has no relationships part', () => {
  const parts = new Map([
    ['_rels/.rels', encode(ROOT_RELS)],
    [
      'xl/workbook.xml',
      encode('<workbook><sheets><sheet name="A" r:id="rId1"/></sheets></workbook>'),
    ],
  ])

  assert.deepEqual(openWorkbook(writeContainer({ parts })).sheets, [])
})
