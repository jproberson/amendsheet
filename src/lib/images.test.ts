import assert from 'node:assert/strict'
import { test } from 'node:test'
import {
  appendAnchors,
  buildDrawing,
  imageContentType,
  imageType,
  pictureAnchor,
} from './images.js'

test('imageType recognises PNG, JPEG and GIF, and nothing else', () => {
  assert.equal(imageType(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d])), 'png')
  assert.equal(imageType(new Uint8Array([0xff, 0xd8, 0xff, 0xe0])), 'jpeg')
  assert.equal(imageType(new Uint8Array([0x47, 0x49, 0x46, 0x38])), 'gif')
  assert.equal(imageType(new Uint8Array([0x00, 0x01, 0x02])), undefined)
  assert.equal(imageContentType('png'), 'image/png')
})

test('pictureAnchor spans the corners and embeds the relationship', () => {
  const xml = pictureAnchor({ column: 0, row: 0 }, { column: 3, row: 10 }, 2, 'rId1')
  assert.match(
    xml,
    /<xdr:from><xdr:col>0<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>0<\/xdr:row>/,
  )
  assert.match(
    xml,
    /<xdr:to><xdr:col>3<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>10<\/xdr:row>/,
  )
  assert.match(xml, /<a:blip[^>]* r:embed="rId1"\/>/)
  assert.match(xml, /<xdr:cNvPr id="2" name="Picture 2"\/>/)
})

test('buildDrawing wraps anchors and appendAnchors adds more before the close', () => {
  const one = pictureAnchor({ column: 0, row: 0 }, { column: 1, row: 1 }, 2, 'rId1')
  const drawing = buildDrawing([one])
  assert.match(drawing, /<xdr:wsDr [^>]*><xdr:twoCellAnchor/)
  assert.match(drawing, /<\/xdr:twoCellAnchor><\/xdr:wsDr>$/)

  const two = pictureAnchor({ column: 2, row: 2 }, { column: 3, row: 3 }, 3, 'rId2')
  const both = appendAnchors(drawing, [two])
  assert.equal(both.match(/<xdr:twoCellAnchor/g)?.length, 2)
  assert.match(both, /id="3"[\s\S]*<\/xdr:wsDr>$/)
})
