/**
 * `code` is what to branch on. Messages are meant to be read by people and may
 * be reworded; the code will not change without a major version.
 */
export type XlsxErrorCode =
  | 'not-a-zip'
  | 'missing-part'
  | 'unreadable-part'
  | 'malformed-xml'
  | 'bad-reference'
  | 'unwritable-value'

export interface XlsxErrorContext {
  readonly part?: string
  readonly reference?: string
  readonly cause?: unknown
}

export class XlsxError extends Error {
  readonly code: XlsxErrorCode
  /** Where in the document the failure was, when that is known. */
  readonly part: string | undefined
  readonly reference: string | undefined

  constructor(code: XlsxErrorCode, message: string, context: XlsxErrorContext = {}) {
    super(message, { cause: context.cause })
    this.name = 'XlsxError'
    this.code = code
    this.part = context.part
    this.reference = context.reference
  }
}
