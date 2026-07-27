import { readXml } from './xml.js'

const escapeXml = (text: string): string =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // XML 1.0 folds a bare CR to LF before the application sees it, so a literal
    // one written raw is lost; the numeric reference survives.
    .replace(/\r/g, '&#13;')

/**
 * Reads a comments part into a map of cell reference to text. A comment's text is
 * rich — runs of differently-formatted spans — and this keeps only the words,
 * joining every `<t>` it holds, since the model reports a comment as a string.
 */
export function readComments(xml: string): ReadonlyMap<string, string> {
  const comments = new Map<string, string>()
  let reference: string | undefined
  let text = ''
  let inText = false
  for (const event of readXml(xml)) {
    if (event.kind === 'open' && event.localName === 'comment') {
      reference = event.attributes.get('ref')
      text = ''
    } else if (event.kind === 'open' && event.localName === 't') {
      inText = true
    } else if (event.kind === 'close' && event.localName === 't') {
      inText = false
    } else if (event.kind === 'text' && inText) {
      text += event.text
    } else if (event.kind === 'close' && event.localName === 'comment' && reference !== undefined) {
      comments.set(reference, text)
      reference = undefined
    }
  }
  return comments
}

const COMMENTS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/**
 * Builds a fresh comments part. One empty author holds every note, since the
 * model carries a note's text but not who wrote it. `xml:space="preserve"` keeps
 * leading and trailing spaces a reader would otherwise trim.
 */
export function buildCommentsPart(entries: ReadonlyMap<string, string>): string {
  const list = [...entries]
    .map(
      ([reference, text]) =>
        `<comment ref="${reference}" authorId="0"><text>` +
        `<t xml:space="preserve">${escapeXml(text)}</t></text></comment>`,
    )
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<comments xmlns="${COMMENTS_NS}"><authors><author/></authors>` +
    `<commentList>${list}</commentList></comments>`
  )
}
