import { decodeText } from './zip.js'
import type { Parts } from './types.js'

/**
 * Features identified by the presence of whole ZIP parts. These are the
 * expensive things a user would be angry to lose: charts, images, pivot
 * tables, macros.
 */
const PART_FEATURES: Array<[string, RegExp]> = [
  ['charts', /^xl\/charts\/chart\d+\.xml$/],
  ['images', /^xl\/media\//],
  ['pivotTables', /^xl\/pivotTables\/pivotTable\d+\.xml$/],
  ['pivotCaches', /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/],
  ['drawings', /^xl\/drawings\/drawing\d+\.xml$/],
  ['comments', /^xl\/(comments|threadedComments)\d*\.xml$/],
  ['vbaMacros', /^xl\/vbaProject\.bin$/],
  ['tables', /^xl\/tables\/table\d+\.xml$/],
  ['theme', /^xl\/theme\/theme\d+\.xml$/],
  ['customXml', /^customXml\//],
]

/**
 * Features identified by markup inside the sheet/workbook/styles XML.
 * Counted by occurrence so we can see partial loss, not just total loss.
 */
const XML_FEATURES: Array<[string, RegExp]> = [
  ['conditionalFormatting', /<conditionalFormatting[\s>]/g],
  ['dataValidation', /<dataValidation[\s>]/g],
  ['autoFilter', /<autoFilter[\s>]/g],
  ['mergedCells', /<mergeCell[\s/>]/g],
  ['hyperlinks', /<hyperlink[\s>]/g],
  ['formulas', /<f[\s>]/g],
  ['frozenPanes', /<pane[\s>]/g],
  ['definedNames', /<definedName[\s>]/g],
  ['colWidths', /<col[\s>]/g],
]

/**
 * Deliberately NOT counted here: entries in the styles.xml registries
 * (numFmt/font/fill/border). Their counts measure the registry, not the
 * document's appearance — a writer that drops an unused format or dedupes
 * two identical fonts has lost nothing. Formatting that a cell actually
 * uses is compared per-cell instead, via the style fingerprint on CellValue.
 */

const XML_SCAN_TARGETS = /^xl\/(workbook\.xml|styles\.xml|worksheets\/.*\.xml)$/

export type FeatureCounts = Map<string, number>

export function scanFeatures(parts: Parts): FeatureCounts {
  const counts: FeatureCounts = new Map()
  const bump = (key: string, by = 1) => counts.set(key, (counts.get(key) ?? 0) + by)

  for (const [path, bytes] of parts) {
    // Directory entries carry no content, and not every writer emits them.
    if (path.endsWith('/')) continue

    for (const [feature, pattern] of PART_FEATURES) {
      if (pattern.test(path)) bump(feature)
    }

    if (!XML_SCAN_TARGETS.test(path)) continue

    const xml = decodeText(bytes)
    for (const [feature, pattern] of XML_FEATURES) {
      const found = xml.match(pattern)
      if (found) bump(feature, found.length)
    }
  }

  return counts
}
