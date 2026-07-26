import assert from 'node:assert/strict'
import { test } from 'node:test'
import { paletteByIndex, readThemeColors, resolveColor } from './theme.js'

const THEME = `<theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><themeElements><clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</clrScheme></themeElements></theme>`

test('readThemeColors maps each scheme slot to its hex, honouring the dk1/lt1 swap', () => {
  const palette = paletteByIndex(readThemeColors(THEME))
  // Index 0 is the light background, index 1 the dark text — the swap.
  assert.equal(palette[0], 'FFFFFF')
  assert.equal(palette[1], '000000')
  assert.equal(palette[2], 'E7E6E6')
  assert.equal(palette[3], '44546A')
  assert.equal(palette[4], '4472C4')
  assert.equal(palette[9], '70AD47')
  assert.equal(palette[10], '0563C1')
  assert.equal(palette[11], '954F72')
})

const palette = paletteByIndex(readThemeColors(THEME))

test('resolveColor returns a plain hex as 8-digit ARGB', () => {
  assert.equal(resolveColor('4472C4', palette), 'FF4472C4')
  assert.equal(resolveColor('#4472c4', palette), 'FF4472C4')
  assert.equal(resolveColor('FF4472C4', palette), 'FF4472C4')
})

test('resolveColor returns a theme colour without tint as its palette hex', () => {
  assert.equal(resolveColor({ theme: 4 }, palette), 'FF4472C4')
})

test('resolveColor darkens with a negative tint and lightens with a positive one', () => {
  // Integer HLS over 0–240: white at tint -0.5 halves luminance to 120, which is
  // 7F on the byte, and black at tint +0.5 raises it from 0 to the same 120.
  assert.equal(resolveColor({ theme: 0, tint: -0.5 }, palette), 'FF7F7F7F')
  assert.equal(resolveColor({ theme: 1, tint: 0.5 }, palette), 'FF7F7F7F')
})

test('resolveColor tints a saturated theme colour close to what Excel shows', () => {
  // accent1 4472C4 darkened 50%. Excel shows 203864; the integer HLS lands within
  // one unit per channel, the model's documented tolerance.
  assert.equal(resolveColor({ theme: 4, tint: -0.499985 }, palette), 'FF203764')
})

test('resolveColor reads the legacy indexed palette', () => {
  assert.equal(resolveColor({ indexed: 10 }, palette), 'FFFF0000')
  assert.equal(resolveColor({ indexed: 8 }, palette), 'FF000000')
})

test('resolveColor returns undefined for a system indexed colour with no fixed hex', () => {
  assert.equal(resolveColor({ indexed: 64 }, palette), undefined)
})

test('resolveColor returns undefined for a theme index the palette does not cover', () => {
  assert.equal(resolveColor({ theme: 99 }, palette), undefined)
})

test('resolveColor tints across the hue and luminance range', () => {
  // Each base has a different channel at its max and sits at a different
  // luminance, so between them they exercise every arm of the HLS conversion.
  assert.equal(resolveColor({ theme: 5, tint: 0.4 }, palette), 'FFF4B084') // ED7D31, red max
  assert.equal(resolveColor({ theme: 9, tint: -0.3 }, palette), 'FF4F7A32') // 70AD47, green max
  assert.equal(resolveColor({ theme: 8, tint: 0.2 }, palette), 'FF7BAEDD') // 5B9BD5, blue max
  assert.equal(resolveColor({ theme: 7, tint: -0.25 }, palette), 'FFBF8F00') // FFC000, high lum
  assert.equal(resolveColor({ theme: 3, tint: 0.6 }, palette), 'FFADBACB') // 44546A, low lum
})

test('resolveColor rejects a malformed hex and reads an 8-digit one', () => {
  assert.equal(resolveColor('12345678', palette), '12345678')
  assert.equal(resolveColor('nothex', palette), undefined)
})

test('resolveColor returns undefined for an indexed colour past the palette', () => {
  assert.equal(resolveColor({ indexed: 200 }, palette), undefined)
})

test('resolveColor treats a zero tint as no tint', () => {
  assert.equal(resolveColor({ theme: 4, tint: 0 }, palette), 'FF4472C4')
})

test('resolveColor tints magenta- and cyan-dominant colours', () => {
  const p = paletteByIndex(
    readThemeColors(
      `<clrScheme><a:accent1><a:srgbClr val="FF00FF"/></a:accent1>` +
        `<a:accent2><a:srgbClr val="00FFFF"/></a:accent2></clrScheme>`,
    ),
  )
  assert.equal(resolveColor({ theme: 4, tint: -0.4 }, p), 'FF990099')
  assert.equal(resolveColor({ theme: 5, tint: 0.4 }, p), 'FF66FFFF')
})

test('readThemeColors drops a slot whose colour has no usable value', () => {
  const colors = readThemeColors(
    `<clrScheme><a:dk1><a:srgbClr/></a:dk1><a:lt1><a:prstClr val="black"/></a:lt1></clrScheme>`,
  )
  assert.equal(colors.get('dk1'), undefined)
  assert.equal(colors.get('lt1'), undefined)
})

test('readThemeColors fills a bare sysClr from its known name and drops the rest', () => {
  const colors = readThemeColors(
    `<clrScheme><a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>` +
      `<a:dk2><a:sysClr val="mysteryColor"/></a:dk2><a:lt2><a:srgbClr val="ZZZZZZ"/></a:lt2></clrScheme>`,
  )
  assert.equal(colors.get('dk1'), '000000')
  assert.equal(colors.get('lt1'), 'FFFFFF')
  assert.equal(colors.get('dk2'), undefined)
  assert.equal(colors.get('lt2'), undefined)
})
