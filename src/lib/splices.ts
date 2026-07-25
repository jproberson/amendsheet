export interface Splice {
  readonly start: number
  readonly end: number
  /** Empty to cut the span out, otherwise what replaces it. */
  readonly text: string
}

/** Applies non-overlapping splices to a string, left to right. */
export function applySplices(text: string, splices: readonly Splice[]): string {
  let out = ''
  let position = 0
  for (const splice of [...splices].sort((a, b) => a.start - b.start)) {
    out += text.slice(position, splice.start) + splice.text
    position = splice.end
  }
  return out + text.slice(position)
}
