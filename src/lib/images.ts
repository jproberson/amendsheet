import { EMPTY_RELATIONSHIPS, type ContainerDraft, withRelationship } from './container-draft.js'
import { withDrawing } from './drawings.js'
import { relationshipsPathFor } from './workbook-parts.js'

/** The picture formats this library embeds, keyed by the file extension Excel
 * expects. Each is recognised from the first bytes of the image, so a caller need
 * not name the type. */
export type ImageType = 'png' | 'jpeg' | 'gif'

const CONTENT_TYPES: Readonly<Record<ImageType, string>> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
}

export const imageContentType = (type: ImageType): string => CONTENT_TYPES[type]

/** The image's format from its magic bytes, or undefined for one this library does
 * not embed. PNG, JPEG and GIF cover what a spreadsheet normally carries. */
export function imageType(bytes: Uint8Array): ImageType | undefined {
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpeg'
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'gif'
  return undefined
}

const SPREADSHEET_DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const DRAWING_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const DRAWING_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

interface Corner {
  readonly column: number
  readonly row: number
}

/** One picture's `twoCellAnchor`, spanning `from` to `to` (both zero-based cell
 * corners), embedding the media through `relationshipId`. `editAs="oneCell"` moves
 * the picture with its cells but keeps its size, the way Excel anchors a pasted one. */
export function pictureAnchor(
  from: Corner,
  to: Corner,
  shapeId: number,
  relationshipId: string,
): string {
  const anchor = (tag: string, corner: Corner): string =>
    `<xdr:${tag}><xdr:col>${corner.column}</xdr:col><xdr:colOff>0</xdr:colOff>` +
    `<xdr:row>${corner.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${tag}>`
  return (
    '<xdr:twoCellAnchor editAs="oneCell">' +
    anchor('from', from) +
    anchor('to', to) +
    '<xdr:pic>' +
    `<xdr:nvPicPr><xdr:cNvPr id="${shapeId}" name="Picture ${shapeId}"/>` +
    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
    `<xdr:blipFill><a:blip xmlns:r="${DRAWING_RELATIONSHIPS}" r:embed="${relationshipId}"/>` +
    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>' +
    '<xdr:clientData/></xdr:twoCellAnchor>'
  )
}

/** A fresh drawing part wrapping the given anchors. */
export function buildDrawing(anchors: readonly string[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<xdr:wsDr xmlns:xdr="${SPREADSHEET_DRAWING}" xmlns:a="${DRAWING_MAIN}">` +
    `${anchors.join('')}</xdr:wsDr>`
  )
}

/** Appends more anchors to an existing drawing, before its closing tag. */
export function appendAnchors(drawingXml: string, anchors: readonly string[]): string {
  const close = drawingXml.lastIndexOf('</xdr:wsDr>')
  if (close === -1) return drawingXml
  return drawingXml.slice(0, close) + anchors.join('') + drawingXml.slice(close)
}

/** A picture pending embedding: its media bytes and the cell corners it spans. */
export interface PendingImage {
  readonly bytes: Uint8Array
  readonly type: ImageType
  readonly from: { readonly column: number; readonly row: number }
  readonly to: { readonly column: number; readonly row: number }
}

const DRAWING_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
const IMAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
export const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml'

/**
 * Embeds this session's images into the draft. Each becomes a media part
 * referenced from the sheet's drawing through a fresh relationship, its picture
 * anchored over the given cells; a sheet with no drawing gets one wired by a
 * `<drawing>`, one that already has a drawing (its own or from another image this
 * session) has the picture appended. Each fresh drawing part and media extension
 * declares its content type on the draft.
 */
export function contributeImages(
  draft: ContainerDraft,
  images: ReadonlyMap<string, readonly PendingImage[]>,
  removedSheets: ReadonlySet<string>,
): void {
  const encoder = new TextEncoder()
  // A media path embeds the image type, so numbers are shared across types —
  // image1.png then image2.jpeg, never two image1s. A per-path freeNumber would
  // hand out image1.jpeg beside image1.png, so this one stays a shared counter.
  let mediaNumber = 0

  for (const [path, imageList] of images) {
    if (removedSheets.has(path) || imageList.length === 0) continue

    const relationshipsPath = relationshipsPathFor(path)
    const existingDrawing = draft.relationshipTarget(path, DRAWING_RELATIONSHIP)

    const drawingPath =
      existingDrawing ??
      `xl/drawings/drawing${draft.freeNumber((n) => `xl/drawings/drawing${n}.xml`)}.xml`
    const drawingRelsPath = relationshipsPathFor(drawingPath)
    let drawingRels = draft.text(drawingRelsPath) ?? EMPTY_RELATIONSHIPS
    const existingDrawingXml = draft.text(drawingPath)
    let shapeId = 1
    for (const match of (existingDrawingXml ?? '').matchAll(/<xdr:cNvPr id="(\d+)"/g)) {
      shapeId = Math.max(shapeId, Number(match[1]))
    }

    const anchors: string[] = []
    for (const image of imageList) {
      do {
        mediaNumber += 1
      } while (draft.has(`xl/media/image${mediaNumber}.${image.type}`))
      const mediaPath = `xl/media/image${mediaNumber}.${image.type}`
      draft.setBytes(mediaPath, image.bytes)
      draft.declareDefault(image.type, imageContentType(image.type))
      const rel = withRelationship(
        drawingRels,
        IMAGE_RELATIONSHIP,
        `../${mediaPath.slice('xl/'.length)}`,
      )
      drawingRels = rel.xml
      anchors.push(pictureAnchor(image.from, image.to, ++shapeId, rel.id))
    }

    draft.setBytes(
      drawingPath,
      encoder.encode(
        existingDrawingXml === undefined
          ? buildDrawing(anchors)
          : appendAnchors(existingDrawingXml, anchors),
      ),
    )
    draft.setBytes(drawingRelsPath, encoder.encode(drawingRels))

    if (existingDrawing === undefined) {
      draft.declareOverride(drawingPath, DRAWING_CONTENT_TYPE)
      const wired = withRelationship(
        draft.text(relationshipsPath),
        DRAWING_RELATIONSHIP,
        `../${drawingPath.slice('xl/'.length)}`,
      )
      draft.setBytes(relationshipsPath, encoder.encode(wired.xml))
      const sheetXml = draft.text(path)
      if (sheetXml !== undefined)
        draft.setBytes(path, encoder.encode(withDrawing(sheetXml, wired.id)))
    }
  }
}
