import type { Color } from './styles-writer.js'
import { readXml } from './xml.js'

const HEX6 = /^[0-9A-Fa-f]{6}$/

/** Fills in a `sysClr` that omits `lastClr`, the two that name a colour. */
const SYSTEM_COLOR_FALLBACK = new Map([
  ['windowText', '000000'],
  ['window', 'FFFFFF'],
])

const colorOfSchemeChild = (
  name: string,
  attributes: ReadonlyMap<string, string>,
): string | undefined => {
  if (name === 'srgbClr') {
    const value = attributes.get('val')
    return value !== undefined && HEX6.test(value) ? value.toUpperCase() : undefined
  }
  if (name === 'sysClr') {
    const last = attributes.get('lastClr')
    if (last !== undefined && HEX6.test(last)) return last.toUpperCase()
    const value = attributes.get('val')
    return value === undefined ? undefined : SYSTEM_COLOR_FALLBACK.get(value)
  }
  return undefined
}

const SCHEME_SLOTS = new Set([
  'dk1',
  'lt1',
  'dk2',
  'lt2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
])

/** Reads `theme1.xml`'s colour scheme into a slot-name to RGB-hex map. Slots the
 * file leaves as an unnamed system colour are simply absent. */
export function readThemeColors(themeXml: string): ReadonlyMap<string, string> {
  const colors = new Map<string, string>()
  let slot: string | undefined
  let inScheme = false
  for (const event of readXml(themeXml)) {
    if (event.kind === 'open' && event.localName === 'clrScheme') {
      inScheme = true
      continue
    }
    if (event.kind === 'close' && event.localName === 'clrScheme') break
    if (!inScheme) continue
    if (event.kind === 'open' && SCHEME_SLOTS.has(event.localName)) {
      slot = event.localName
      continue
    }
    if (event.kind === 'close' && event.localName === slot) {
      slot = undefined
      continue
    }
    if (event.kind === 'open' && slot !== undefined) {
      const hex = colorOfSchemeChild(event.localName, event.attributes)
      if (hex !== undefined) colors.set(slot, hex)
    }
  }
  return colors
}

// The `theme` index in a cell's colour reads the scheme in a different order than
// the file stores it: index 0 is the light background and index 1 the dark text,
// so `dk1`/`lt1` (and `dk2`/`lt2`) are swapped. A SpreadsheetML quirk, not a bug.
const THEME_INDEX_ORDER: readonly string[] = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
]

export function paletteByIndex(
  colors: ReadonlyMap<string, string>,
): readonly (string | undefined)[] {
  return THEME_INDEX_ORDER.map((slot) => colors.get(slot))
}

// The default indexed palette (ECMA-376 §18.8.27). Indices 0–7 repeat 8–15;
// 64 and 65 are the system foreground and background, which carry no fixed hex.
const INDEXED_PALETTE: readonly (string | undefined)[] = [
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '000000',
  'FFFFFF',
  'FF0000',
  '00FF00',
  '0000FF',
  'FFFF00',
  'FF00FF',
  '00FFFF',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'C0C0C0',
  '808080',
  '9999FF',
  '993366',
  'FFFFCC',
  'CCFFFF',
  '660066',
  'FF8080',
  '0066CC',
  'CCCCFF',
  '000080',
  'FF00FF',
  'FFFF00',
  '00FFFF',
  '800080',
  '800000',
  '008080',
  '0000FF',
  '00CCFF',
  'CCFFFF',
  'CCFFCC',
  'FFFF99',
  '99CCFF',
  'FF99CC',
  'CC99FF',
  'FFCC99',
  '3366FF',
  '33CCCC',
  '99CC00',
  'FFCC00',
  'FF9900',
  'FF6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
]

const toHex = (value: number): string =>
  Math.min(255, Math.max(0, value)).toString(16).padStart(2, '0').toUpperCase()

// Excel resolves a tint in the Win32 HLS colour space, which uses integer
// arithmetic with a 0–240 range rather than floats over 0–1. Matching it to the
// byte means matching that model, down to the round-half terms; a float HSL is
// off by one in a channel often enough to show. From MS KB "How To Convert
// Colors Between RGB and HLS", the algorithm ECMA-376 §18.8.19 points a tint at.
const HLSMAX = 240
const RGBMAX = 255

/** Integer division truncating toward zero, as the C source does. Every operand
 * below is non-negative, so this is a floor. */
const div = (numerator: number, denominator: number): number => Math.floor(numerator / denominator)

interface Hls {
  readonly hue: number
  readonly luminance: number
  readonly saturation: number
}

const rgbToHls = (red: number, green: number, blue: number): Hls => {
  const cMax = Math.max(red, green, blue)
  const cMin = Math.min(red, green, blue)
  const luminance = div((cMax + cMin) * HLSMAX + RGBMAX, 2 * RGBMAX)
  if (cMax === cMin) return { hue: 0, luminance, saturation: 0 }
  const span = cMax - cMin
  const saturation =
    luminance <= HLSMAX / 2
      ? div(span * HLSMAX + div(cMax + cMin, 2), cMax + cMin)
      : div(span * HLSMAX + div(2 * RGBMAX - cMax - cMin, 2), 2 * RGBMAX - cMax - cMin)
  const sixth = div(HLSMAX, 6)
  const half = div(span, 2)
  const rDelta = div((cMax - red) * sixth + half, span)
  const gDelta = div((cMax - green) * sixth + half, span)
  const bDelta = div((cMax - blue) * sixth + half, span)
  let hue: number
  if (red === cMax) hue = bDelta - gDelta
  else if (green === cMax) hue = div(HLSMAX, 3) + rDelta - bDelta
  else hue = div(2 * HLSMAX, 3) + gDelta - rDelta
  // The KB source also clamps hue above HLSMAX here, but the integer deltas cap
  // this expression near 200, so that arm cannot run; `hueToChannel` still wraps.
  if (hue < 0) hue += HLSMAX
  return { hue, luminance, saturation }
}

const hueToChannel = (magic1: number, magic2: number, hueIn: number): number => {
  let hue = hueIn
  if (hue < 0) hue += HLSMAX
  if (hue > HLSMAX) hue -= HLSMAX
  const sixth = div(HLSMAX, 6)
  const twelfth = div(HLSMAX, 12)
  if (hue < sixth) return magic1 + div((magic2 - magic1) * hue + twelfth, sixth)
  if (hue < div(HLSMAX, 2)) return magic2
  if (hue < div(2 * HLSMAX, 3))
    return magic1 + div((magic2 - magic1) * (div(2 * HLSMAX, 3) - hue) + twelfth, sixth)
  return magic1
}

const hlsToRgb = ({ hue, luminance, saturation }: Hls): string => {
  if (saturation === 0) {
    const grey = toHex(div(luminance * RGBMAX, HLSMAX))
    return `${grey}${grey}${grey}`
  }
  const magic2 =
    luminance <= HLSMAX / 2
      ? div(luminance * (HLSMAX + saturation) + div(HLSMAX, 2), HLSMAX)
      : luminance + saturation - div(luminance * saturation + div(HLSMAX, 2), HLSMAX)
  const magic1 = 2 * luminance - magic2
  const third = div(HLSMAX, 3)
  const channel = (hueOffset: number): string =>
    toHex(div(hueToChannel(magic1, magic2, hue + hueOffset) * RGBMAX + div(HLSMAX, 2), HLSMAX))
  return `${channel(third)}${channel(0)}${channel(-third)}`
}

/** A negative tint scales luminance toward 0, a positive one toward its max. The
 * hue and saturation are left as they were. */
const applyTint = (rgb6: string, tint: number): string => {
  const red = Number.parseInt(rgb6.slice(0, 2), 16)
  const green = Number.parseInt(rgb6.slice(2, 4), 16)
  const blue = Number.parseInt(rgb6.slice(4, 6), 16)
  const hls = rgbToHls(red, green, blue)
  const luminance = Math.round(
    tint < 0 ? hls.luminance * (1 + tint) : hls.luminance * (1 - tint) + HLSMAX * tint,
  )
  return hlsToRgb({ hue: hls.hue, luminance, saturation: hls.saturation })
}

const normalizeHex = (color: string): string | undefined => {
  const hex = color.startsWith('#') ? color.slice(1) : color
  if (HEX6.test(hex)) return `FF${hex.toUpperCase()}`
  if (/^[0-9A-Fa-f]{8}$/.test(hex)) return hex.toUpperCase()
  return undefined
}

/**
 * Resolves a stored colour reference to the 8-digit ARGB hex it displays as, or
 * `undefined` when it names something with no fixed value — a theme slot the
 * palette does not cover, or a system indexed colour. `palette` is the workbook's
 * theme colours by index, from `paletteByIndex(readThemeColors(...))`.
 *
 * A plain hex and an indexed colour resolve exactly. A tinted theme colour is
 * computed in the same integer HLS space Excel uses, and lands within one unit
 * per channel of what Excel displays — imperceptible, but not always bit-exact.
 */
export function resolveColor(
  color: Color,
  palette: readonly (string | undefined)[],
): string | undefined {
  if (typeof color === 'string') return normalizeHex(color)
  if ('theme' in color) {
    const base = palette[color.theme]
    if (base === undefined) return undefined
    const rgb = color.tint === undefined || color.tint === 0 ? base : applyTint(base, color.tint)
    return `FF${rgb}`
  }
  const base = INDEXED_PALETTE[color.indexed]
  return base === undefined ? undefined : `FF${base}`
}
