import assert from 'node:assert/strict'
import { test } from 'node:test'
import { applyLineShifts } from './apply-line-shifts.js'
import type { Container } from './container.js'
import { createContainerDraft } from './container-draft.js'

const encode = (text: string) => new TextEncoder().encode(text)
const decode = (bytes: Uint8Array | null | undefined) =>
  bytes ? new TextDecoder().decode(bytes) : undefined

const containerOf = (parts: Record<string, string>): Container => ({
  parts: new Map(Object.entries(parts).map(([path, text]) => [path, encode(text)])),
})

const insertRowsAt = (editedSheet: string, at: number, count: number) => ({
  axis: 'row' as const,
  at,
  delta: count,
  editedSheet,
  onCurrentSheet: true,
})

test('with no line ops the pending names pass through unchanged', () => {
  const changes = new Map<string, Uint8Array | null>()
  const draft = createContainerDraft(containerOf({}), changes)
  const pending = new Map([['Total', 'Sheet1!$A$1']])
  const names = applyLineShifts(draft, containerOf({}), changes, {
    sheets: [],
    addedSheetPaths: [],
    removed: new Set(),
    lineOps: [],
    fileNames: new Map(),
    pendingNames: pending,
    date1904: false,
  })
  assert.equal(names, pending)
  assert.equal(changes.size, 0)
})

test('a defined name pointing past the insert moves down', () => {
  const changes = new Map<string, Uint8Array | null>()
  const container = containerOf({})
  const draft = createContainerDraft(container, changes)
  const names = applyLineShifts(draft, container, changes, {
    sheets: [],
    addedSheetPaths: [],
    removed: new Set(),
    lineOps: [{ path: 'xl/worksheets/sheet1.xml', spec: insertRowsAt('Sheet1', 3, 2) }],
    fileNames: new Map([['Anchor', 'Sheet1!$A$5']]),
    pendingNames: new Map(),
    date1904: false,
  })
  assert.equal(names.get('Anchor'), 'Sheet1!$A$7')
})

test('an inserted row shifts the cells on the edited sheet', () => {
  const path = 'xl/worksheets/sheet1.xml'
  const sheet =
    '<worksheet><sheetData><row r="5"><c r="A5"><v>1</v></c></row></sheetData></worksheet>'
  const changes = new Map<string, Uint8Array | null>()
  const container = containerOf({ [path]: sheet })
  const draft = createContainerDraft(container, changes)
  applyLineShifts(draft, container, changes, {
    sheets: [{ path, name: 'Sheet1' }],
    addedSheetPaths: [],
    removed: new Set(),
    lineOps: [{ path, spec: insertRowsAt('Sheet1', 3, 2) }],
    fileNames: new Map(),
    pendingNames: new Map(),
    date1904: false,
  })
  const written = decode(changes.get(path))
  assert.ok(written?.includes('r="7"'), 'row 5 should have shifted to row 7')
  assert.ok(written?.includes('r="A7"'), 'cell A5 should have shifted to A7')
})
