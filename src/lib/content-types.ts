import { XlsxError } from './errors.js'

/** Declares a part in the content types with an Override, once. */
export function withContentTypeOverride(
  xml: string,
  partName: string,
  contentType: string,
): string {
  if (xml.includes(`PartName="/${partName}"`)) return xml
  const override = `<Override PartName="/${partName}" ContentType="${contentType}"/>`
  const close = xml.indexOf('</Types>')
  if (close === -1) {
    throw new XlsxError('invalid-content', 'Content types part is malformed', {
      part: '[Content_Types].xml',
    })
  }
  return xml.slice(0, close) + override + xml.slice(close)
}

/** Declares a file extension in the content types with a Default, once. */
export function withContentTypeDefault(
  xml: string,
  extension: string,
  contentType: string,
): string {
  if (new RegExp(`<Default Extension="${extension}"`).test(xml)) return xml
  const close = xml.indexOf('>')
  if (close === -1 || !xml.startsWith('<')) {
    throw new XlsxError('invalid-content', 'Content types part is malformed', {
      part: '[Content_Types].xml',
    })
  }
  // A Default sits among the other Defaults at the top of the Types element.
  const typesOpen = xml.indexOf('<Types')
  const after = xml.indexOf('>', typesOpen) + 1
  const element = `<Default Extension="${extension}" ContentType="${contentType}"/>`
  return xml.slice(0, after) + element + xml.slice(after)
}
