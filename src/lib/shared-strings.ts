import { readXml } from './xml.js'

/**
 * One string may be split across formatting runs, and `rPh` runs hold phonetic
 * guides that look like text but are not part of the value.
 */
export function readSharedStrings(xml: string): readonly string[] {
  const strings: string[] = []

  let current: string[] | null = null
  let inPhonetic = false
  let inText = false

  for (const event of readXml(xml)) {
    if (event.kind === 'open') {
      if (event.name === 'si') current = []
      else if (event.name === 'rPh') inPhonetic = true
      else if (event.name === 't' && !event.selfClosing) inText = !inPhonetic
      continue
    }

    if (event.kind === 'text') {
      if (inText && current !== null) current.push(event.text)
      continue
    }

    if (event.name === 't') inText = false
    else if (event.name === 'rPh') inPhonetic = false
    else if (event.name === 'si' && current !== null) {
      strings.push(current.join(''))
      current = null
    }
  }

  return strings
}
