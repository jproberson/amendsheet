import { parseReference } from './reference.js'
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

/**
 * Builds a fresh comments part. One empty author holds every note, since the
 * model carries a note's text but not who wrote it. `xml:space="preserve"` keeps
 * leading and trailing spaces a reader would otherwise trim. The text goes in an
 * `<r>` run: a bare `<t>` is valid CT_Rst, but Excel and other readers write and
 * expect a run, and some drop the text of a runless comment.
 */
export function buildCommentsPart(entries: ReadonlyMap<string, string>): string {
  const list = [...entries]
    .map(
      ([reference, text]) =>
        `<comment ref="${reference}" authorId="0"><text>` +
        `<r><t xml:space="preserve">${escapeXml(text)}</t></r></text></comment>`,
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<comments xmlns="${COMMENTS_NS}"><authors><author/></authors>` +
    `<commentList>${list}</commentList></comments>`
  )
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
 * Builds the legacy VML drawing that gives each note its box. The text lives in
 * the comments part; the box's shape, position and size live here, and without
 * it Excel stores the note but draws nothing. Each shape borrows the shared
 * `_x0000_t202` textbox type, stays hidden until the cell is hovered, and
 * anchors to its cell's zero-based row and column. The anchor's finer offsets
 * are the defaults Excel writes; it recomputes them when it lays the note out.
 */
export function buildVmlDrawing(references: readonly string[]): string {
  const shapes = references
    .map((reference, index) => {
      const { row, column } = parseReference(reference)
      const cellRow = row - 1
      const cellColumn = column - 1
      const anchor = [cellColumn + 1, 15, cellRow, 2, cellColumn + 3, 15, cellRow + 4, 4].join(', ')
      return (
        `<v:shape id="_x0000_s${1025 + index}" type="#_x0000_t202" ` +
        'style="position:absolute;margin-left:60pt;margin-top:1.5pt;' +
        `width:108pt;height:60pt;z-index:${index + 1};visibility:hidden" ` +
        'fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/>' +
        '<v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>' +
        '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
        `<x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
        `<x:Row>${cellRow}</x:Row><x:Column>${cellColumn}</x:Column></x:ClientData></v:shape>`
      )
    })
    .join('')
  return `${VML_HEADER}${shapes}</xml>`
}
