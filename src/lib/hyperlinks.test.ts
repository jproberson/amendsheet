import assert from 'node:assert/strict'
import { test } from 'node:test'

import { withHyperlinkRelationships, withHyperlinks, writeSheetHyperlinks } from './hyperlinks.js'

const worksheet = (body: string) =>
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</worksheet>`

test('withHyperlinks with nothing to write leaves the sheet as it is', () => {
  assert.equal(
    withHyperlinks('<worksheet><sheetData/></worksheet>', []),
    '<worksheet><sheetData/></worksheet>',
  )
})

test('withHyperlinks adds a fresh element after the sheet body', () => {
  const out = withHyperlinks(worksheet('<sheetData/>'), [
    { reference: 'A1', location: 'Sheet2!B2' },
  ])
  assert.match(
    out,
    /<sheetData\/><hyperlinks><hyperlink ref="A1" location="Sheet2!B2"\/><\/hyperlinks><\/worksheet>/,
  )
})

test('withHyperlinks places the element before a following page element', () => {
  const out = withHyperlinks(worksheet('<sheetData/><pageMargins left="0"/>'), [
    { reference: 'A1', location: 'B2' },
  ])
  assert.match(out, /<\/sheetData>|<sheetData\/>/)
  assert.match(out, /<hyperlinks><hyperlink ref="A1" location="B2"\/><\/hyperlinks><pageMargins/)
})

test('withHyperlinks writes an external link with its relationship id and declares xmlns:r', () => {
  const out = withHyperlinks(worksheet('<sheetData/>'), [
    { reference: 'A1', relationshipId: 'rId1', tooltip: 'Go <there>' },
  ])
  assert.match(
    out,
    /<worksheet [^>]*xmlns:r="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships"/,
  )
  assert.match(out, /<hyperlink ref="A1" r:id="rId1" tooltip="Go &lt;there&gt;"\/>/)
})

test('withHyperlinks merges into an existing element and replaces a link on the same cell', () => {
  const sheet = worksheet(
    '<sheetData/><hyperlinks><hyperlink ref="A1" location="old"/><hyperlink ref="C3" location="keep"/></hyperlinks>',
  )
  const out = withHyperlinks(sheet, [
    { reference: 'A1', location: 'new' },
    { reference: 'B2', location: 'add' },
  ])
  assert.doesNotMatch(out, /location="old"/)
  assert.match(out, /<hyperlink ref="C3" location="keep"\/>/)
  assert.match(out, /<hyperlink ref="A1" location="new"\/>/)
  assert.match(out, /<hyperlink ref="B2" location="add"\/>/)
})

test('withHyperlinks fills a self-closing element', () => {
  const out = withHyperlinks(worksheet('<sheetData/><hyperlinks/>'), [
    { reference: 'A1', location: 'B2' },
  ])
  assert.match(out, /<hyperlinks><hyperlink ref="A1" location="B2"\/><\/hyperlinks>/)
})

test('withHyperlinks keeps the sheet prefix and an existing xmlns:r', () => {
  const sheet = '<x:worksheet xmlns:x="x" xmlns:r="r"><x:sheetData/></x:worksheet>'
  const out = withHyperlinks(sheet, [{ reference: 'A1', relationshipId: 'rId3' }])
  assert.match(out, /<x:sheetData\/><x:hyperlinks><x:hyperlink ref="A1" r:id="rId3"\/>/)
  assert.equal(out.match(/xmlns:r=/g)?.length, 1)
})

test('withHyperlinks appends to a fragment that has no worksheet root', () => {
  assert.equal(
    withHyperlinks('<sheetData/>', [{ reference: 'A1', location: 'B2' }]),
    '<sheetData/><hyperlinks><hyperlink ref="A1" location="B2"/></hyperlinks>',
  )
  // An external link on a rootless fragment cannot declare xmlns:r, and does not try.
  assert.equal(
    withHyperlinks('<sheetData/>', [{ reference: 'A1', relationshipId: 'rId1' }]),
    '<sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>',
  )
})

test('withHyperlinks skips a hyperlink with no ref when replacing', () => {
  const sheet = worksheet('<sheetData/><hyperlinks><hyperlink location="orphan"/></hyperlinks>')
  const out = withHyperlinks(sheet, [{ reference: 'A1', location: 'B2' }])
  assert.match(out, /<hyperlink location="orphan"\/><hyperlink ref="A1" location="B2"\/>/)
})

test('withHyperlinkRelationships creates the part and appends to it', () => {
  const created = withHyperlinkRelationships(undefined, [
    { id: 'rId1', url: 'https://a.example/x' },
  ])
  assert.match(created, /<Relationships xmlns="[^"]*relationships">/)
  assert.match(
    created,
    /<Relationship Id="rId1" Type="[^"]*hyperlink" Target="https:\/\/a.example\/x" TargetMode="External"\/>/,
  )
  const appended = withHyperlinkRelationships(
    '<Relationships xmlns="x"><Relationship Id="rId1" Type="t" Target="d"/></Relationships>',
    [{ id: 'rId2', url: 'https://b.example/y?q=1&z=2' }],
  )
  assert.match(appended, /Id="rId1"/)
  assert.match(
    appended,
    /Id="rId2" Type="[^"]*hyperlink" Target="https:\/\/b.example\/y\?q=1&amp;z=2"/,
  )
})

test('writeSheetHyperlinks gives an external link a fresh id past the sheet rels', () => {
  const written = writeSheetHyperlinks(
    worksheet('<sheetData/>'),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="x" Target="y"/></Relationships>',
    new Map([['A1', { url: 'https://example.com' }]]),
  )
  assert.match(written.sheetXml, /<hyperlink ref="A1" r:id="rId4"\/>/)
  assert.match(
    written.relsXml ?? '',
    /Id="rId4"[^>]*Target="https:\/\/example.com"[^>]*TargetMode="External"/,
  )
})

test('writeSheetHyperlinks writes an internal location inline and needs no rels', () => {
  const written = writeSheetHyperlinks(
    worksheet('<sheetData/>'),
    undefined,
    new Map([['B2', { location: 'Sheet2!C3', tooltip: 'see C3' }]]),
  )
  assert.match(written.sheetXml, /<hyperlink ref="B2" location="Sheet2!C3" tooltip="see C3"\/>/)
  assert.equal(written.relsXml, undefined)
})
