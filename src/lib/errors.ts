/**
 * Every failure this library raises. `part` names the location inside the
 * document when one is known, because "Invalid row number in model" is not
 * something a caller can act on.
 */
export class XlsxError extends Error {
  readonly part: string | undefined

  constructor(message: string, options: { part?: string; cause?: unknown } = {}) {
    super(message, { cause: options.cause })
    this.name = 'XlsxError'
    this.part = options.part
  }
}
