import assert from 'node:assert/strict'
import { test } from 'node:test'
import { shiftPivotCacheSource, shiftPivotLocation } from './pivots.js'
import type { ShiftSpec } from './shift.js'

const rowSpec = (at: number, delta: number, editedSheet = 'Data'): ShiftSpec => ({
  axis: 'row',
  at,
  delta,
  editedSheet,
  onCurrentSheet: true,
})

test('shiftPivotLocation moves the location range and leaves the offsets', () => {
  const xml =
    '<pivotTableDefinition><location ref="A10:D14" firstDataRow="2"/></pivotTableDefinition>'
  const out = shiftPivotLocation(xml, rowSpec(1, 1))
  assert.match(out, /<location ref="A11:D15" firstDataRow="2"\/>/)
})

test('shiftPivotCacheSource shifts a source on the edited sheet, matching case', () => {
  const xml =
    '<pivotCacheDefinition><cacheSource><worksheetSource ref="A1:C4" sheet="data"/></cacheSource></pivotCacheDefinition>'
  assert.match(
    shiftPivotCacheSource(xml, rowSpec(1, 2, 'DATA')),
    /<worksheetSource ref="A3:C6" sheet="data"\/>/,
  )
})

test('shiftPivotCacheSource leaves a source on another sheet alone', () => {
  const xml =
    '<pivotCacheDefinition><worksheetSource ref="A1:C4" sheet="Other"/></pivotCacheDefinition>'
  assert.equal(shiftPivotCacheSource(xml, rowSpec(1, 2, 'Data')), xml)
})

test('shiftPivotCacheSource leaves a name-based source alone', () => {
  const xml = '<pivotCacheDefinition><worksheetSource name="SalesTable"/></pivotCacheDefinition>'
  assert.equal(shiftPivotCacheSource(xml, rowSpec(1, 2, 'Data')), xml)
})
