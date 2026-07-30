import type { ShiftSpec } from './shift.js'
import { readXml } from './xml.js'

// A corner at or past the edit moves; a deletion clamps one inside the removed
// band to the surviving edge — the top up to the line after it, the bottom down
// to the line before. Values are the one-based lines the anchor corners sit on.
const shiftLow = (value: number, spec: ShiftSpec): number => {
  if (value < spec.at) return value
  if (spec.delta > 0) return value + spec.delta
  return value < spec.at - spec.delta ? spec.at : value + spec.delta
}
const shiftHigh = (value: number, spec: ShiftSpec): number => {
  if (value < spec.at) return value
  if (spec.delta > 0) return value + spec.delta
  return value < spec.at - spec.delta ? spec.at - 1 : value + spec.delta
}

interface Corner {
  start: number
  end: number
  value: number
}

interface Splice {
  start: number
  end: number
  text: string
}

/**
 * Moves each drawing's cell anchors with the rows or columns an edit shifts under
 * them. A `twoCellAnchor` moves its top-left and bottom-right corners on the edited
 * axis, shrinking when a deletion clips it and dropping when a deletion takes every
 * line it spanned; a `oneCellAnchor` moves only its top-left. Corners are zero-based
 * cell indices, so each is raised to a one-based line for the shift.
 */
export function shiftDrawing(drawingXml: string, spec: ShiftSpec): string {
  const axis = spec.axis === 'row' ? 'row' : 'col'
  const splices: Splice[] = []
  let anchorStart = -1
  let side: 'from' | 'to' | undefined
  let capturing = false
  let text = ''
  let fieldStart = -1
  let from: Corner | undefined
  let to: Corner | undefined
  for (const event of readXml(drawingXml)) {
    if (event.kind === 'open' && event.localName.endsWith('CellAnchor')) {
      anchorStart = event.start
      from = undefined
      to = undefined
    } else if (event.kind === 'open' && (event.localName === 'from' || event.localName === 'to')) {
      side = event.localName
    } else if (event.kind === 'close' && (event.localName === 'from' || event.localName === 'to')) {
      side = undefined
    } else if (event.kind === 'open' && event.localName === axis && side !== undefined) {
      capturing = true
      text = ''
      fieldStart = event.end
    } else if (event.kind === 'text' && capturing) {
      text += event.text
    } else if (event.kind === 'close' && event.localName === axis && capturing) {
      const corner = { start: fieldStart, end: event.start, value: Number(text) }
      if (side === 'from') from = corner
      else to = corner
      capturing = false
    } else if (
      event.kind === 'close' &&
      event.localName.endsWith('CellAnchor') &&
      anchorStart !== -1
    ) {
      const newFrom = from === undefined ? undefined : shiftLow(from.value + 1, spec) - 1
      const newTo = to === undefined ? undefined : shiftHigh(to.value + 1, spec) - 1
      if (newFrom !== undefined && newTo !== undefined && newTo < newFrom) {
        const anchorEnd = drawingXml.indexOf('>', event.start) + 1
        splices.push({ start: anchorStart, end: anchorEnd, text: '' })
      } else {
        if (from !== undefined && newFrom !== from.value)
          splices.push({ start: from.start, end: from.end, text: String(newFrom) })
        if (to !== undefined && newTo !== to.value)
          splices.push({ start: to.start, end: to.end, text: String(newTo) })
      }
      anchorStart = -1
    }
  }
  let xml = drawingXml
  for (const splice of splices.sort((a, b) => b.start - a.start)) {
    xml = xml.slice(0, splice.start) + splice.text + xml.slice(splice.end)
  }
  return xml
}
