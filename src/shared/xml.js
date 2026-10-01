/**
 * Lightweight, robust XML utilities for OOXML processing.
 * Handles entity escaping, tag querying, text extraction, and XML node manipulation.
 */

export function escapeXml(str) {
  if (typeof str === 'undefined' || str === null) return ''
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

export function unescapeXml(str) {
  if (typeof str === 'undefined' || str === null) return ''
  return String(str)
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, '>')
    .replace(/&lt;/g, '<')
    .replace(/&amp;/g, '&')
}

/**
 * Extract plain text content from XML string by stripping tags.
 * @param {string} xml
 * @returns {string}
 */
export function extractTextFromXml(xml) {
  if (!xml) return ''
  let text = xml
    .replace(/<w:br\s*\/?>/gi, '\n')
    .replace(/<w:cr\s*\/?>/gi, '\n')
    .replace(/<w:tab\s*\/?>/gi, '\t')
    .replace(/<[^>]+>/g, '')
  return unescapeXml(text)
}

/**
 * Extract all matching XML elements by tag name.
 * Uses exact tag boundary matching so `<w:tbl>` does not match `<w:tblPr>`.
 * @param {string} xml
 * @param {string} tagName - e.g. 'w:p', 'w:tr', 'w:tc', 'w:t', 'w:tbl'
 * @returns {Array<{outerXml: string, innerXml: string, index: number}>}
 */
export function extractElements(xml, tagName) {
  const results = []
  const escapedTag = tagName.replace(':', '\\:')
  // Lookahead ensures tag boundary: followed by space, >, or /
  const openRegex = new RegExp(`<${escapedTag}(?=[\\s>/])([^>]*)?>`, 'gi')
  const closeTag = `</${tagName}>`

  let match
  while ((match = openRegex.exec(xml)) !== null) {
    const startIndex = match.index
    const openTagLen = match[0].length

    // Handle self-closing tag
    if (match[0].endsWith('/>')) {
      results.push({
        outerXml: match[0],
        innerXml: '',
        index: startIndex
      })
      continue
    }

    // Find matching close tag with support for nested tags of same name
    let depth = 1
    let searchPos = startIndex + openTagLen
    let endPos = -1

    const nestedOpenRegex = new RegExp(`<${escapedTag}(?=[\\s>/])`, 'gi')

    while (depth > 0 && searchPos < xml.length) {
      nestedOpenRegex.lastIndex = searchPos
      const nextOpenMatch = nestedOpenRegex.exec(xml)
      const nextOpen = nextOpenMatch ? nextOpenMatch.index : -1
      const nextClose = xml.indexOf(closeTag, searchPos)

      if (nextClose === -1) {
        break
      }

      if (nextOpen !== -1 && nextOpen < nextClose) {
        const nextOpenEnd = xml.indexOf('>', nextOpen)
        if (nextOpenEnd !== -1 && xml[nextOpenEnd - 1] === '/') {
          searchPos = nextOpenEnd + 1
        } else {
          depth++
          searchPos = nextOpen + tagName.length + 1
        }
      } else {
        depth--
        if (depth === 0) {
          endPos = nextClose + closeTag.length
          break
        }
        searchPos = nextClose + closeTag.length
      }
    }

    if (endPos !== -1) {
      const outerXml = xml.substring(startIndex, endPos)
      const innerXml = xml.substring(startIndex + openTagLen, endPos - closeTag.length)
      results.push({
        outerXml,
        innerXml,
        index: startIndex
      })
      openRegex.lastIndex = endPos
    }
  }

  return results
}

/**
 * Extract an attribute value from an XML opening tag.
 *
 * The attribute name is anchored at an XML name boundary. Without that anchor
 * a substring match returns the wrong value whenever one attribute name ends
 * with another: R7 writes `<col customWidth="1" min="1" max="1" width="14"/>`,
 * so an unanchored lookup of `width` yields "1" (from `customWidth`) and a
 * lookup of `ht` yields "1" (from `customHeight`). Silent wrong values like
 * that corrupt documents, which is why the boundary is enforced here rather
 * than in each caller.
 *
 * @param {string} tagXml
 * @param {string} attrName - e.g. "r", "w:val", "customWidth"
 * @returns {string|null}
 */
export function getAttribute(tagXml, attrName) {
  if (!tagXml || !attrName) return null
  const escaped = attrName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  // A preceding character may be the start of the string, whitespace, or the
  // opening angle bracket; anything else would make this part of a longer name.
  const regex = new RegExp(`(?:^|[\\s<])${escaped}\\s*=\\s*["']([^"']*)["']`)
  const match = tagXml.match(regex)
  return match ? match[1] : null
}
