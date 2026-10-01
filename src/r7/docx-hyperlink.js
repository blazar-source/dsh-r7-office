/**
 * DOCX hyperlinks.
 *
 * A hyperlink is two coupled facts:
 *
 *   - `word/document.xml` carries `<w:hyperlink r:id="rIdN">` wrapping the
 *     visible runs;
 *   - `word/_rels/document.xml.rels` carries the matching relationship of type
 *     `…/hyperlink`, marked `TargetMode="External"` because the target is a
 *     URL rather than a part inside the package.
 *
 * Editing the text must not touch the relationship, and inserting a link must
 * not disturb the ones already there — including their order, because a
 * relationship id is referenced by the document and never by position.
 */

import {
  REL_TYPES,
  addRelationship,
  childElements,
  escapeAttr as escapeAttrValue,
  parseRelationships
} from './docx-parts.js'
import { appendToBodyEnd } from './docx-sections.js'
import { extractElements, extractTextFromXml, getAttribute } from '../shared/xml.js'
import { isInsideElement } from './docx-parts.js'
import { normalizeAlignment } from './docx-styles.js'

/** The character style Word and R7 apply to hyperlink runs. */
export const HYPERLINK_STYLE_ID = 'Hyperlink'

/**
 * Every hyperlink in the document, in document order.
 *
 * @param {string} docXml
 * @param {Map<string, object>|Array<object>} relationships
 * @returns {Array<{index: number, paragraphIndex: number, text: string, url: string|null, anchor: string|null, relId: string|null, tooltip: string|null, target: string|null, runStyle: string|null}>}
 */
export function listHyperlinks(docXml, relationships) {
  const byId = relationships instanceof Map
    ? relationships
    : new Map((relationships || []).map(rel => [rel.id, rel]))
  const result = []
  const paragraphs = extractElements(docXml, 'w:p')

  for (let i = 0; i < paragraphs.length; i++) {
    for (const element of childElements(paragraphs[i].outerXml)) {
      if (element.tag !== 'w:hyperlink') continue
      const relId = getAttribute(element.xml, 'r:id')
      const anchor = getAttribute(element.xml, 'w:anchor')
      const rel = relId ? byId.get(relId) : null
      const runStyle = getAttribute(extractElements(element.xml, 'w:rStyle')[0]?.outerXml ?? '', 'w:val')
      result.push({
        index: result.length,
        paragraphIndex: i,
        text: extractTextFromXml(element.xml),
        relId,
        anchor,
        url: rel ? rel.target : null,
        target: rel ? rel.target : null,
        tooltip: getAttribute(element.xml, 'w:tooltip'),
        runStyle
      })
    }
  }

  return result
}

/**
 * Find or create an external relationship for a URL.
 *
 * Reusing an existing relationship keeps the relationship part stable across
 * repeated insertions of the same link instead of growing it by one entry per
 * call.
 *
 * @param {string} relsXml
 * @param {string} url
 * @returns {{xml: string, id: string, created: boolean}}
 */
export function ensureExternalRelationship(relsXml, url) {
  if (!relsXml) {
    throw new Error('Invalid DOCX: word/_rels/document.xml.rels is missing')
  }
  const existing = parseRelationships(relsXml).list
    .find(rel => rel.type === REL_TYPES.hyperlink && rel.target === url && (rel.targetMode || '').toLowerCase() === 'external')
  if (existing) return { xml: relsXml, id: existing.id, created: false }

  const added = addRelationship(relsXml, {
    type: REL_TYPES.hyperlink,
    target: url,
    targetMode: 'External'
  })
  return { xml: added.xml, id: added.id, created: true }
}

/**
 * Build a `<w:hyperlink>` element.
 *
 * @param {object} spec
 * @param {string} [spec.text]
 * @param {string} [spec.relId] - external target relationship
 * @param {string} [spec.anchor] - internal bookmark/anchor name
 * @param {string} [spec.tooltip]
 * @param {string} [spec.style] - character style id (default `Hyperlink`)
 * @param {boolean} [spec.bold]
 * @param {boolean} [spec.italic]
 * @param {number} [spec.size]
 * @param {string} [spec.color]
 * @returns {string}
 */
export function buildHyperlinkXml(spec = {}) {
  const text = spec.text === undefined || spec.text === null ? '' : String(spec.text)
  const rPrParts = []
  const style = spec.style === undefined ? HYPERLINK_STYLE_ID : spec.style
  if (style) rPrParts.push(`<w:rStyle w:val="${escapeAttrValue(style)}"/>`)
  if (spec.bold) rPrParts.push('<w:b/>')
  if (spec.italic) rPrParts.push('<w:i/>')
  if (spec.size !== undefined && spec.size !== null) {
    rPrParts.push(`<w:sz w:val="${Math.round(Number(spec.size) * 2)}"/>`)
  }
  if (spec.color) rPrParts.push(`<w:color w:val="${escapeAttrValue(String(spec.color).replace(/^#/, ''))}"/>`)

  const rPr = rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : ''
  const run = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`

  const attributes = []
  if (spec.relId) attributes.push(`r:id="${escapeAttrValue(spec.relId)}"`)
  if (spec.anchor) attributes.push(`w:anchor="${escapeAttrValue(spec.anchor)}"`)
  if (spec.tooltip) attributes.push(`w:tooltip="${escapeAttrValue(spec.tooltip)}"`)

  if (!spec.relId && !spec.anchor) {
    throw new Error('A hyperlink needs either a relationship id (external URL) or an anchor')
  }

  return `<w:hyperlink ${attributes.join(' ')}>${run}</w:hyperlink>`
}

function escapeXmlText(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * Insert a hyperlink into a paragraph or as a new paragraph.
 *
 * @param {string} docXml
 * @param {object} options
 * @param {string} options.hyperlinkXml - the element built by {@link buildHyperlinkXml}
 * @param {string} [options.text]
 * @param {string} [options.url] - external target (creates/reuses a relationship)
 * @param {string} [options.anchor] - internal anchor
 * @param {string} [options.tooltip]
 * @param {number|'end'} [options.paragraphIndex='end'] - the paragraph the link joins
 * @param {boolean} [options.newParagraph=false] - insert the link as its own paragraph
 * @param {string} [options.alignment] - alignment for the new paragraph
 * @returns {{docXml: string, paragraphIndex: number, hyperlinkXml: string}}
 */
export function insertHyperlink(docXml, options) {
  const hyperlinkXml = options.hyperlinkXml
  if (!hyperlinkXml) throw new Error('insertHyperlink requires options.hyperlinkXml')

  const paragraphs = extractElements(docXml, 'w:p')
  const requested = options.paragraphIndex ?? 'end'
  let targetIndex
  if (requested === 'end') targetIndex = paragraphs.length - 1
  else if (requested === 'start') targetIndex = 0
  else {
    targetIndex = Number(requested)
    if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= paragraphs.length) {
      throw new Error(`Paragraph index out of range for a hyperlink: ${requested} (total: ${paragraphs.length})`)
    }
  }

  if (options.newParagraph) {
    const alignment = normalizeAlignment(options.alignment)
    const pPr = alignment ? `<w:pPr><w:jc w:val="${alignment}"/></w:pPr>` : ''
    const paragraph = `<w:p>${pPr}${hyperlinkXml}</w:p>`
    const anchorParagraph = paragraphs[targetIndex]

    // A paragraph that lives inside a table cell cannot host a body-level
    // paragraph: falling back to the end of the body keeps the document valid
    // instead of dropping the link into the last cell of the last table.
    const insideTable = anchorParagraph
      ? isInsideElement(docXml, 'w:tbl', anchorParagraph.index)
      : false
    if (!anchorParagraph || insideTable || requested === 'end') {
      const appended = appendToBodyEnd(docXml, paragraph)
      return {
        docXml: appended,
        paragraphIndex: extractElements(appended, 'w:p').length - 1,
        hyperlinkXml
      }
    }

    const at = anchorParagraph.index + anchorParagraph.outerXml.length
    return {
      docXml: docXml.slice(0, at) + paragraph + docXml.slice(at),
      paragraphIndex: targetIndex + 1,
      hyperlinkXml
    }
  }

  const target = paragraphs[targetIndex]
  if (!target) throw new Error('Cannot insert a hyperlink: the document has no paragraphs')
  const end = target.outerXml.lastIndexOf('</w:p>')
  const updated = `${target.outerXml.slice(0, end)}${hyperlinkXml}${target.outerXml.slice(end)}`
  return {
    docXml: docXml.slice(0, target.index) + updated + docXml.slice(target.index + target.outerXml.length),
    paragraphIndex: targetIndex,
    hyperlinkXml
  }
}

/**
 * Replace the visible text of an existing hyperlink, preserving its
 * relationship and its run formatting.
 *
 * @param {string} docXml
 * @param {object} options
 * @param {number} [options.index] - hyperlink index from {@link listHyperlinks}
 * @param {number} [options.paragraphIndex] - first hyperlink in that paragraph
 * @param {string} options.text
 * @returns {{docXml: string, index: number, paragraphIndex: number, url: string|null}}
 */
export function setHyperlinkText(docXml, options, relationships) {
  const links = listHyperlinks(docXml, relationships)
  let target = null
  if (options.index !== undefined && options.index !== null) {
    target = links.find(link => link.index === Number(options.index)) || null
  } else if (options.paragraphIndex !== undefined && options.paragraphIndex !== null) {
    target = links.find(link => link.paragraphIndex === Number(options.paragraphIndex)) || null
  } else if (links.length === 1) {
    target = links[0]
  } else {
    throw new Error('setHyperlinkText needs an index or a paragraphIndex')
  }
  if (!target) {
    throw new Error(`Hyperlink not found (index=${options.index ?? 'n/a'}, paragraphIndex=${options.paragraphIndex ?? 'n/a'})`)
  }

  const paragraphs = extractElements(docXml, 'w:p')
  const paragraph = paragraphs[target.paragraphIndex]
  const element = childElements(paragraph.outerXml).find(entry => entry.tag === 'w:hyperlink')
  if (!element) throw new Error('Hyperlink element vanished between listing and editing')

  const runs = extractElements(element.xml, 'w:r')
  let updated
  if (runs.length === 0) {
    const close = element.xml.lastIndexOf('</w:hyperlink>')
    updated = `${element.xml.slice(0, close)}<w:r><w:t xml:space="preserve">${escapeXmlText(options.text ?? '')}</w:t></w:r>${element.xml.slice(close)}`
  } else {
    updated = element.xml
    // Replace the first run's text and empty any others, keeping their rPr.
    for (let i = runs.length - 1; i >= 0; i--) {
      const run = runs[i]
      const tElements = extractElements(run.outerXml, 'w:t')
      if (tElements.length === 0) continue
      let newRun = run.outerXml
      for (let j = tElements.length - 1; j >= 0; j--) {
        const t = tElements[j]
        const value = (i === 0 && j === 0)
          ? `<w:t xml:space="preserve">${escapeXmlText(options.text ?? '')}</w:t>`
          : '<w:t></w:t>'
        newRun = newRun.slice(0, t.index) + value + newRun.slice(t.index + t.outerXml.length)
      }
      const offset = run.index
      updated = updated.slice(0, offset) + newRun + updated.slice(offset + run.outerXml.length)
    }
  }

  const updatedParagraph = paragraph.outerXml.slice(0, element.start)
    + updated
    + paragraph.outerXml.slice(element.end)
  const xml = docXml.slice(0, paragraph.index) + updatedParagraph + docXml.slice(paragraph.index + paragraph.outerXml.length)

  return { docXml: xml, index: target.index, paragraphIndex: target.paragraphIndex, url: target.url }
}

/**
 * Remove a hyperlink, keeping its visible text as plain runs.
 * @param {string} docXml
 * @param {object} options
 * @param {number} options.index
 * @param {Map<string, object>|Array<object>} relationships
 * @returns {{docXml: string, index: number}}
 */
export function removeHyperlink(docXml, options, relationships) {
  const links = listHyperlinks(docXml, relationships)
  const target = links.find(link => link.index === Number(options.index))
  if (!target) throw new Error(`Hyperlink index out of range: ${options.index} (found ${links.length})`)

  const paragraphs = extractElements(docXml, 'w:p')
  const paragraph = paragraphs[target.paragraphIndex]
  const element = childElements(paragraph.outerXml).find(entry => entry.tag === 'w:hyperlink')
  if (!element) throw new Error('Hyperlink element vanished between listing and editing')

  const runs = extractElements(element.xml, 'w:r').map(run => run.outerXml).join('')
  const updatedParagraph = paragraph.outerXml.slice(0, element.start) + runs + paragraph.outerXml.slice(element.end)
  const xml = docXml.slice(0, paragraph.index) + updatedParagraph + docXml.slice(paragraph.index + paragraph.outerXml.length)
  return { docXml: xml, index: target.index }
}
