import assert from 'node:assert/strict'
import { test } from 'node:test'
import { drawingHasUnshiftableFrame, shiftDrawing } from './drawings.js'
import type { ShiftSpec } from './shift.js'

const spec = (axis: 'row' | 'column', at: number, delta: number): ShiftSpec => ({
  axis,
  at,
  delta,
  editedSheet: 'Sheet1',
  onCurrentSheet: true,
})

const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const twoCell = (from: [number, number], to: [number, number]) =>
  `<xdr:twoCellAnchor><xdr:from><xdr:col>${from[1]}</xdr:col><xdr:colOff>0</xdr:colOff>` +
  `<xdr:row>${from[0]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
  `<xdr:to><xdr:col>${to[1]}</xdr:col><xdr:colOff>0</xdr:colOff>` +
  `<xdr:row>${to[0]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
  `<xdr:pic/><xdr:clientData/></xdr:twoCellAnchor>`
const wsDr = (body: string) => `<xdr:wsDr xmlns:xdr="${XDR}">${body}</xdr:wsDr>`

test('drawingHasUnshiftableFrame is true for a non-chart frame, false for a chart or none', () => {
  const frame = (uri: string) =>
    `<xdr:twoCellAnchor><xdr:graphicFrame><a:graphic><a:graphicData uri="${uri}"/></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor>`
  assert.equal(drawingHasUnshiftableFrame(wsDr(twoCell([0, 0], [3, 3]))), false) // just a picture
  assert.equal(
    drawingHasUnshiftableFrame(
      wsDr(frame('http://schemas.openxmlformats.org/drawingml/2006/chart')),
    ),
    false, // a chart is shiftable now
  )
  assert.equal(
    drawingHasUnshiftableFrame(
      wsDr(frame('http://schemas.openxmlformats.org/drawingml/2006/diagram')),
    ),
    true, // a diagram is not
  )
})

test('shiftDrawing moves a column anchor and leaves rows alone', () => {
  const out = shiftDrawing(wsDr(twoCell([1, 2], [4, 5])), spec('column', 1, 2))
  assert.match(out, /<xdr:from><xdr:col>4<\/xdr:col>.*<xdr:row>1<\/xdr:row>/)
  assert.match(out, /<xdr:to><xdr:col>7<\/xdr:col>.*<xdr:row>4<\/xdr:row>/)
})

test('shiftDrawing shrinks an anchor a deletion clips and drops one it swallows', () => {
  // Rows 2-4 (0-based) span 1-based rows 3-5; deleting rows 3-5 swallows it.
  const swallowed = shiftDrawing(wsDr(twoCell([2, 0], [4, 3])), spec('row', 3, -3))
  assert.doesNotMatch(swallowed, /<xdr:twoCellAnchor>/)
  // Rows 2-6 only overlap the top of the deletion, so it shrinks rather than goes.
  const clipped = shiftDrawing(wsDr(twoCell([2, 0], [6, 3])), spec('row', 3, -3))
  assert.match(clipped, /<xdr:from>.*<xdr:row>2<\/xdr:row>/)
  assert.match(clipped, /<xdr:to>.*<xdr:row>3<\/xdr:row>/) // bottom pulled up to row 4 -> 3
})

test('shiftDrawing moves a oneCellAnchor by its single corner', () => {
  const one =
    '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>' +
    '<xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
    '<xdr:ext cx="100" cy="100"/><xdr:pic/></xdr:oneCellAnchor>'
  const out = shiftDrawing(wsDr(one), spec('row', 2, 1))
  assert.match(out, /<xdr:row>6<\/xdr:row>/)
})
