/** `part` names where in the document the failure was, when that is known. */
export class XlsxError extends Error {
  readonly part: string | undefined

  constructor(message: string, options: { part?: string; cause?: unknown } = {}) {
    super(message, { cause: options.cause })
    this.name = 'XlsxError'
    this.part = options.part
  }
}
