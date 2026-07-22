/**
 * `code` is what to branch on. Messages are meant to be read by people and may
 * be reworded.
 *
 * Two codes mean the caller passed something the library cannot use, and are
 * the ones worth answering with a different value: `bad-reference` and
 * `unwritable-value`. Every other code is about the file, or the runtime
 * reading it, not a value you passed.
 *
 * The union is open. An existing code will not change without a major version,
 * but new ones arrive in minor releases, because there is no knowing today
 * every way an xlsx can be malformed. A switch over this needs a default.
 */
export type XlsxErrorCode =
  | 'not-a-zip'
  | 'missing-part'
  | 'unreadable-part'
  | 'malformed-xml'
  /** Well formed xml that says something no reader can honour. */
  | 'invalid-content'
  /** A part decompresses to more bytes than this runtime can hold at once. */
  | 'part-too-large'
  | 'bad-reference'
  | 'unwritable-value'
  | (string & {})

/**
 * Where a failure happened, as tightly as the throw site knows it. A cell
 * failure carries all three of `part`, `sheet` and `reference`; a part-level
 * failure carries `part`; only a whole-file failure with nowhere finer to point
 * (a non-zip, a low-level parse offset already in the message) carries just
 * `cause`. The argument is required so no throw can silently drop the location.
 */
export interface XlsxErrorContext {
  readonly part?: string
  /** The worksheet name, when a cell or sheet is the locus. */
  readonly sheet?: string
  readonly reference?: string
  readonly cause?: unknown
}

export class XlsxError extends Error {
  readonly code: XlsxErrorCode
  /** Where in the document the failure was, as tightly as it was known. */
  readonly part: string | undefined
  readonly sheet: string | undefined
  readonly reference: string | undefined

  constructor(code: XlsxErrorCode, message: string, context: XlsxErrorContext) {
    super(message, { cause: context.cause })
    this.name = 'XlsxError'
    this.code = code
    this.part = context.part
    this.sheet = context.sheet
    this.reference = context.reference
  }
}
