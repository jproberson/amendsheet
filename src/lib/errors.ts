/**
 * `code` is what to branch on. Messages are meant to be read by people and may
 * be reworded.
 *
 * Two codes mean the caller passed something the library cannot use, and are
 * the ones worth answering with a different value: `bad-reference` and
 * `unwritable-value`. Every other code means the file is at fault.
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
  | 'bad-reference'
  | 'unwritable-value'
  | (string & {})

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
