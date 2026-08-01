import { type ContainerDraft, withRelationship } from './container-draft.js'
import { withLegacyDrawing } from './drawings.js'
import { XlsxError } from './errors.js'
import { formatReference, parseReference } from './reference.js'
import { type ShiftSpec, shiftLine } from './shift.js'
import { relationshipsPathFor } from './workbook-parts.js'
import { readXml } from './xml.js'

const escapeXml = (text: string): string =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // XML 1.0 folds a bare CR to LF before the application sees it, so a literal
    // one written raw is lost; the numeric reference survives.
    .replace(/\r/g, '&#13;')

/**
 * Reads a comments part into a map of cell reference to text. A comment's text is
 * rich — runs of differently-formatted spans — and this keeps only the words,
 * joining every `<t>` it holds, since the model reports a comment as a string.
 */
export function readComments(xml: string): ReadonlyMap<string, string> {
  const comments = new Map<string, string>()
  let reference: string | undefined
  let text = ''
  let inText = false
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'comment') {
      reference = event.attributes.get('ref')
      text = ''
    } else if (event.kind === 'open' && event.localName === 't') {
      inText = true
    } else if (event.kind === 'close' && event.localName === 't') {
      inText = false
    } else if (event.kind === 'text' && inText) {
      text += event.text
    } else if (event.kind === 'close' && event.localName === 'comment' && reference !== undefined) {
      comments.set(reference, text)
      reference = undefined
    }
  }
  return comments
}

const COMMENTS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

// The text goes in an `<r>` run: a bare `<t>` is valid CT_Rst, but Excel and
// other readers write and expect a run, and some drop the text of a runless one.
const commentElement = (reference: string, text: string): string =>
  `<comment ref="${reference}" authorId="0"><text>` +
  `<r><t xml:space="preserve">${escapeXml(text)}</t></r></text></comment>`

/**
 * Builds a fresh comments part. One empty author holds every note, since the
 * model carries a note's text but not who wrote it. `xml:space="preserve"` keeps
 * leading and trailing spaces a reader would otherwise trim.
 */
export function buildCommentsPart(entries: ReadonlyMap<string, string>): string {
  const list = [...entries].map(([reference, text]) => commentElement(reference, text)).join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<comments xmlns="${COMMENTS_NS}"><authors><author/></authors>` +
    `<commentList>${list}</commentList></comments>`
  )
}

/** Removes the `<comment>` for a cell from a comments part, leaving the rest — and
 * the (now perhaps empty) part — as it was. A ref the part lacks is ignored. A cell
 * carries at most one comment, so the first match is the only one. */
export function withoutComment(commentsXml: string, reference: string): string {
  let start = -1
  let matching = false
  for (const event of readXml(commentsXml)) {
    if (event.kind === 'open' && event.localName === 'comment') {
      start = event.start
      matching = event.attributes.get('ref') === reference
    } else if (event.kind === 'close' && event.localName === 'comment' && matching) {
      const end = commentsXml.indexOf('>', event.start) + 1
      return commentsXml.slice(0, start) + commentsXml.slice(end)
    }
  }
  return commentsXml
}

/** Removes the note `<v:shape>` anchored at a cell's zero-based row and column from
 * a legacy drawing, leaving shapes that are not notes, or note other cells, alone.
 * A cell has at most one note shape, so the first match is the only one. */
export function withoutNoteShape(vmlXml: string, cellRow: number, cellColumn: number): string {
  let start = -1
  let row: number | undefined
  let column: number | undefined
  let capture: 'row' | 'column' | undefined
  let text = ''
  for (const event of readXml(vmlXml)) {
    if (event.kind === 'open' && event.localName === 'shape') {
      start = event.start
      row = undefined
      column = undefined
    } else if (event.kind === 'open' && event.localName === 'Row') {
      capture = 'row'
      text = ''
    } else if (event.kind === 'open' && event.localName === 'Column') {
      capture = 'column'
      text = ''
    } else if (event.kind === 'text' && capture !== undefined) {
      text += event.text
    } else if (event.kind === 'close' && event.localName === 'Row') {
      row = Number(text)
      capture = undefined
    } else if (event.kind === 'close' && event.localName === 'Column') {
      column = Number(text)
      capture = undefined
    } else if (event.kind === 'close' && event.localName === 'shape' && start !== -1) {
      if (row === cellRow && column === cellColumn) {
        const end = vmlXml.indexOf('>', event.start) + 1
        return vmlXml.slice(0, start) + vmlXml.slice(end)
      }
      start = -1
    }
  }
  return vmlXml
}

interface Splice {
  readonly start: number
  readonly end: number
  readonly text: string
}

const applySplices = (xml: string, splices: Splice[]): string => {
  let out = xml
  for (const splice of splices.sort((a, b) => b.start - a.start)) {
    out = out.slice(0, splice.start) + splice.text + out.slice(splice.end)
  }
  return out
}

const setRefAttribute = (tag: string, reference: string): string =>
  tag.replace(
    /(\sref\s*=\s*)("[^"]*"|'[^']*')/,
    (_, head: string, quoted: string) => `${head}${quoted[0]}${reference}${quoted[0]}`,
  )

type RefShift = { kind: 'same' } | { kind: 'move'; reference: string } | { kind: 'drop' }

// Where a comment's one cell lands under a row or column edit: unmoved, moved to
// a new cell, or dropped because the edit removed the cell it sat on.
const shiftedRef = (reference: string, spec: ShiftSpec): RefShift => {
  let cell: { row: number; column: number }
  try {
    cell = parseReference(reference)
  } catch {
    return { kind: 'same' }
  }
  const shifted = shiftLine(spec.axis === 'row' ? cell.row : cell.column, spec)
  if (shifted === undefined) return { kind: 'drop' }
  const next = formatReference(
    spec.axis === 'row'
      ? { row: shifted, column: cell.column }
      : { row: cell.row, column: shifted },
  )
  return next === reference ? { kind: 'same' } : { kind: 'move', reference: next }
}

/**
 * Moves each note in a comments part with the row or column its cell sits on, and
 * drops one whose cell a deletion removed. The text and the rest of the element
 * are left byte for byte; only the `ref` moves, or the whole `<comment>` goes.
 */
export function shiftComments(commentsXml: string, spec: ShiftSpec): string {
  const splices: Splice[] = []
  let commentStart = -1
  let openStart = -1
  let openEnd = -1
  let reference: string | undefined
  for (const event of readXml(commentsXml)) {
    if (event.kind === 'open' && event.localName === 'comment') {
      commentStart = event.start
      openStart = event.start
      openEnd = event.end
      reference = event.attributes.get('ref')
    } else if (event.kind === 'close' && event.localName === 'comment' && reference !== undefined) {
      const closeEnd = commentsXml.indexOf('>', event.start) + 1
      const shift = shiftedRef(reference, spec)
      if (shift.kind === 'drop') {
        splices.push({ start: commentStart, end: closeEnd, text: '' })
      } else if (shift.kind === 'move') {
        const rewritten = setRefAttribute(commentsXml.slice(openStart, openEnd), shift.reference)
        splices.push({ start: openStart, end: openEnd, text: rewritten })
      }
      reference = undefined
    }
  }
  return applySplices(commentsXml, splices)
}

// Shifts the row (indices 2 and 6) or column (0 and 4) corners of a note box's
// eight-number anchor by the same amount its cell moved, keeping any spacing. A
// corner that is not a number, or an anchor that is not eight fields, is left be.
const shiftAnchor = (anchor: string, axis: 'row' | 'column', delta: number): string => {
  const parts = anchor.split(',')
  if (parts.length !== 8) return anchor
  const corners = axis === 'row' ? new Set([2, 6]) : new Set([0, 4])
  return parts
    .map((part, index) => {
      if (!corners.has(index)) return part
      const value = Number(part.trim())
      if (!Number.isFinite(value)) return part
      const lead = part.slice(0, part.length - part.trimStart().length)
      return lead + String(value + delta)
    })
    .join(',')
}

interface Field {
  start: number
  end: number
  value: number
}

/**
 * Moves each note's box in a legacy drawing with the row or column its cell sits
 * on — the authoritative `x:Row`/`x:Column` anchor and the box's own corners — and
 * drops a box whose cell a deletion removed. A shape that names no cell is left be.
 */
export function shiftNoteShapes(vmlXml: string, spec: ShiftSpec): string {
  const splices: Splice[] = []
  let shapeStart = -1
  let capture: 'row' | 'column' | 'anchor' | undefined
  let text = ''
  let row: Field | undefined
  let column: Field | undefined
  let anchor: { start: number; end: number; text: string } | undefined
  for (const event of readXml(vmlXml)) {
    if (event.kind === 'open' && event.localName === 'shape') {
      shapeStart = event.start
      row = undefined
      column = undefined
      anchor = undefined
    } else if (event.kind === 'open' && event.localName === 'Row') {
      capture = 'row'
      text = ''
      row = { start: event.end, end: -1, value: 0 }
    } else if (event.kind === 'open' && event.localName === 'Column') {
      capture = 'column'
      text = ''
      column = { start: event.end, end: -1, value: 0 }
    } else if (event.kind === 'open' && event.localName === 'Anchor') {
      capture = 'anchor'
      text = ''
      anchor = { start: event.end, end: -1, text: '' }
    } else if (event.kind === 'text' && capture !== undefined) {
      text += event.text
    } else if (event.kind === 'close' && event.localName === 'Row' && row !== undefined) {
      row.end = event.start
      row.value = Number(text)
      capture = undefined
    } else if (event.kind === 'close' && event.localName === 'Column' && column !== undefined) {
      column.end = event.start
      column.value = Number(text)
      capture = undefined
    } else if (event.kind === 'close' && event.localName === 'Anchor' && anchor !== undefined) {
      anchor.end = event.start
      anchor.text = text
      capture = undefined
    } else if (event.kind === 'close' && event.localName === 'shape' && shapeStart !== -1) {
      if (row !== undefined && column !== undefined) {
        const field = spec.axis === 'row' ? row : column
        const shifted = shiftLine(field.value + 1, spec)
        if (shifted === undefined) {
          const shapeEnd = vmlXml.indexOf('>', event.start) + 1
          splices.push({ start: shapeStart, end: shapeEnd, text: '' })
        } else {
          const delta = shifted - 1 - field.value
          if (delta !== 0) {
            splices.push({ start: field.start, end: field.end, text: String(field.value + delta) })
            if (anchor !== undefined && anchor.end !== -1) {
              const moved = shiftAnchor(anchor.text, spec.axis, delta)
              splices.push({ start: anchor.start, end: anchor.end, text: moved })
            }
          }
        }
      }
      shapeStart = -1
    }
  }
  return applySplices(vmlXml, splices)
}

const VML_HEADER =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" ' +
  'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>' +
  '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'

/**
 * One note's box: a shape borrowing the shared `_x0000_t202` textbox type,
 * hidden until the cell is hovered, anchored to the cell's zero-based row and
 * column. The anchor's finer offsets are the defaults Excel writes; it recomputes
 * them when it lays the note out. The shape id and z-index are passed in so a
 * fresh drawing and one being appended to both stay collision-free.
 */
function noteShape(reference: string, shapeId: number, zIndex: number): string {
  const { row, column } = parseReference(reference)
  const cellRow = row - 1
  const cellColumn = column - 1
  const anchor = [cellColumn + 1, 15, cellRow, 2, cellColumn + 3, 15, cellRow + 4, 4].join(', ')
  return (
    `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" ` +
    'style="position:absolute;margin-left:60pt;margin-top:1.5pt;' +
    `width:108pt;height:60pt;z-index:${zIndex};visibility:hidden" ` +
    'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/>' +
    '<v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
    '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>' +
    '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
    `<x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
    `<x:Row>${cellRow}</x:Row><x:Column>${cellColumn}</x:Column></x:ClientData></v:shape>`
  )
}

/**
 * Builds the legacy VML drawing that gives each note its box. The text lives in
 * the comments part; the box's shape, position and size live here, and without
 * it Excel stores the note but draws nothing.
 */
export function buildVmlDrawing(references: readonly string[]): string {
  const shapes = references.map((reference, index) => noteShape(reference, 1025 + index, index + 1))
  return `${VML_HEADER}${shapes.join('')}</xml>`
}

/**
 * Splices more note shapes into an existing drawing, keeping its bytes and giving
 * each new shape an id and z-index past the highest already there so nothing
 * collides. Excel's shape-id block holds 1024 ids, plenty for the notes a splice
 * adds, so the layout's id map is left as it is.
 */
export function appendVmlShapes(existingXml: string, references: readonly string[]): string {
  let maxShapeId = 1024
  for (const match of existingXml.matchAll(/_x0000_s(\d+)/g))
    maxShapeId = Math.max(maxShapeId, Number(match[1]))
  let maxZIndex = 0
  for (const match of existingXml.matchAll(/z-index:(\d+)/g))
    maxZIndex = Math.max(maxZIndex, Number(match[1]))
  const shapes = references.map((reference, index) =>
    noteShape(reference, maxShapeId + 1 + index, maxZIndex + 1 + index),
  )
  const close = existingXml.lastIndexOf('</xml>')
  if (close === -1) {
    throw new XlsxError('invalid-content', 'A legacy drawing part is malformed', {})
  }
  return existingXml.slice(0, close) + shapes.join('') + existingXml.slice(close)
}

/**
 * Splices more comments into an existing part, keeping its bytes so the rich text
 * the notes already hold survives untouched. New notes take the first author, the
 * one every comments part carries.
 */
export function appendCommentsPart(
  existingXml: string,
  entries: ReadonlyMap<string, string>,
): string {
  const additions = [...entries]
    .map(([reference, text]) => commentElement(reference, text))
    .join('')
  const close = existingXml.indexOf('</commentList>')
  if (close === -1) {
    throw new XlsxError('invalid-content', 'A comments part is malformed', {})
  }
  return existingXml.slice(0, close) + additions + existingXml.slice(close)
}

export const COMMENTS_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
export const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
export const VML_DRAWING_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
export const VML_DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing'

/**
 * Writes this session's comment adds and removals into the draft. Comments live
 * in two parts wired to the sheet and declared in the content types: the comments
 * part holds the text, a legacy VML drawing holds the box Excel draws for each
 * note. A sheet with none yet opens fresh parts; one that already has some has the
 * new notes and shapes spliced in, so its existing rich text survives. Removal
 * runs after the adds, on whatever they just wrote. Each fresh part declares its
 * content type on the draft.
 */
export function contributeComments(
  draft: ContainerDraft,
  comments: ReadonlyMap<string, ReadonlyMap<string, string>>,
  removedComments: ReadonlyMap<string, ReadonlySet<string>>,
  removedSheets: ReadonlySet<string>,
): void {
  const encoder = new TextEncoder()

  for (const [path, notes] of comments) {
    if (removedSheets.has(path) || notes.size === 0) continue

    const relationshipsPath = relationshipsPathFor(path)
    const currentRels = draft.text(relationshipsPath)
    let relsXml = currentRels

    const existingComments = draft.relationshipTarget(path, COMMENTS_RELATIONSHIP)
    if (existingComments === undefined) {
      const commentsPath = `xl/comments${draft.freeNumber((n) => `xl/comments${n}.xml`)}.xml`
      draft.setBytes(commentsPath, encoder.encode(buildCommentsPart(notes)))
      draft.declareOverride(commentsPath, COMMENTS_CONTENT_TYPE)
      relsXml = withRelationship(
        relsXml,
        COMMENTS_RELATIONSHIP,
        `../${commentsPath.slice('xl/'.length)}`,
      ).xml
    } else {
      const existing = draft.original(existingComments) ?? ''
      draft.setBytes(existingComments, encoder.encode(appendCommentsPart(existing, notes)))
    }

    // A sheet's legacy drawing may hold shapes other than notes (form controls),
    // so an existing one is appended to rather than replaced; only a sheet with
    // none needs a drawing authored and pointed at by a fresh <legacyDrawing>.
    const existingVml = draft.relationshipTarget(path, VML_DRAWING_RELATIONSHIP)
    if (existingVml === undefined) {
      const vmlDrawingPath = `xl/drawings/vmlDrawing${draft.freeNumber(
        (n) => `xl/drawings/vmlDrawing${n}.vml`,
      )}.vml`
      draft.setBytes(vmlDrawingPath, encoder.encode(buildVmlDrawing([...notes.keys()])))
      draft.declareOverride(vmlDrawingPath, VML_DRAWING_CONTENT_TYPE)
      const withVml = withRelationship(
        relsXml,
        VML_DRAWING_RELATIONSHIP,
        `../${vmlDrawingPath.slice('xl/'.length)}`,
      )
      relsXml = withVml.xml
      const sheetXml = draft.text(path)
      if (sheetXml !== undefined) {
        draft.setBytes(path, encoder.encode(withLegacyDrawing(sheetXml, withVml.id)))
      }
    } else {
      const existing = draft.original(existingVml) ?? ''
      draft.setBytes(existingVml, encoder.encode(appendVmlShapes(existing, [...notes.keys()])))
    }

    if (relsXml !== currentRels) draft.setBytes(relationshipsPath, encoder.encode(relsXml ?? ''))
  }

  for (const [path, refs] of removedComments) {
    if (removedSheets.has(path)) continue
    const commentsPath = draft.relationshipTarget(path, COMMENTS_RELATIONSHIP)
    const commentsXml = commentsPath === undefined ? undefined : draft.text(commentsPath)
    if (commentsPath !== undefined && commentsXml !== undefined) {
      let updated = commentsXml
      for (const ref of refs) updated = withoutComment(updated, ref)
      if (updated !== commentsXml) draft.setBytes(commentsPath, encoder.encode(updated))
    }
    const vmlPath = draft.relationshipTarget(path, VML_DRAWING_RELATIONSHIP)
    const vmlXml = vmlPath === undefined ? undefined : draft.text(vmlPath)
    if (vmlPath !== undefined && vmlXml !== undefined) {
      let updated = vmlXml
      for (const ref of refs) {
        const { row, column } = parseReference(ref)
        updated = withoutNoteShape(updated, row - 1, column - 1)
      }
      if (updated !== vmlXml) draft.setBytes(vmlPath, encoder.encode(updated))
    }
  }
}
