/**
 * DOCX page setup and sections.
 *
 * A section's properties live in one `<w:sectPr>` element, and there are
 * exactly two places it can be:
 *
 *   - inside a paragraph's `<w:pPr>` — a **section break**: the paragraph ends
 *     the section described by that element;
 *   - as the last child of `<w:body>` — the **final section**, which has no
 *     terminating paragraph.
 *
 * Both forms are handled everywhere in this module, and an edit to one section
 * is applied as a targeted upsert of the individual properties (`w:pgSz`,
 * `w:pgMar`, `w:cols`, …) rather than by replacing the element wholesale. That
 * is what keeps a document's existing headers, footers, columns, page-number
 * restart and document grid alive when only the orientation is changed.
 *
 * Schema order inside `<w:sectPr>` is significant — Word and R7 both validate
 * it — so children are inserted at their schema position, not appended.
 */

import { extractElements, getAttribute, extractTextFromXml } from '../shared/xml.js'
import { twipsToCm, cmToTwips, pointsToTwips } from './docx-styles.js'
import {
  childElements,
  elementParts,
  insertChildOrdered,
  removeChildElements,
  upsertChildOrdered
} from './docx-parts.js'

/** Child order inside `<w:sectPr>`, per the WordprocessingML schema. */
export const SECTPR_ORDER = [
  'w:headerReference',
  'w:footerReference',
  'w:footnotePr',
  'w:endnotePr',
  'w:type',
  'w:pgSz',
  'w:pgMar',
  'w:paperSrc',
  'w:pgBorders',
  'w:lnNumType',
  'w:pgNumType',
  'w:cols',
  'w:formProt',
  'w:vAlign',
  'w:noEndnote',
  'w:titlePg',
  'w:textDirection',
  'w:bidi',
  'w:rtlGutter',
  'w:docGrid',
  'w:printerSettings',
  'w:sectPrChange'
]

/** A4 in twips, the fallback used when a section omits its page size. */
export const A4_TWIPS = { width: 11906, height: 16838 }

/** Section break types. */
export const SECTION_TYPES = new Set(['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'])

/** Page-number formats accepted by `<w:pgNumType w:fmt="…">`. */
export const PAGE_NUMBER_FORMATS = new Set([
  'decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter', 'decimalZero'
])

/**
 * Locate every `<w:sectPr>` in a document body, in document order.
 *
 * @param {string} docXml
 * @returns {Array<{index: number, kind: 'paragraph'|'final', paragraphIndex: number|null, outerXml: string, innerXml: string, position: number}>}
 */
export function collectSections(docXml) {
  const sections = []
  const paragraphs = extractElements(docXml, 'w:p')

  for (let i = 0; i < paragraphs.length; i++) {
    const pXml = paragraphs[i].outerXml
    for (const sectPr of extractElements(pXml, 'w:sectPr')) {
      sections.push({
        index: -1,
        kind: 'paragraph',
        paragraphIndex: i,
        outerXml: sectPr.outerXml,
        innerXml: sectPr.innerXml,
        position: paragraphs[i].index + sectPr.index
      })
    }
  }

  // The body-level section is the last one, and it is a *direct child* of
  // `<w:body>`: a `<w:sectPr>` nested inside a paragraph is a section break,
  // not the final section. Scanning the body's inner XML by tag would find the
  // break paragraph's element first and mistake it for the final one.
  const bodySections = extractElements(docXml, 'w:body')
  const bodyElement = bodySections[0] ?? null
  const finalSectPr = bodyElement
    ? childElements(bodyElement.outerXml).find(element => element.tag === 'w:sectPr') ?? null
    : null
  if (finalSectPr && bodyElement) {
    sections.push({
      index: -1,
      kind: 'final',
      paragraphIndex: null,
      outerXml: finalSectPr.xml,
      innerXml: elementParts(finalSectPr.xml)?.inner ?? '',
      position: bodyElement.index + finalSectPr.start
    })
  }

  sections.sort((a, b) => a.position - b.position)
  sections.forEach((section, idx) => { section.index = idx })
  return sections
}

/**
 * Normalized description of one `<w:sectPr>` element.
 *
 * @param {string} sectPrXml
 * @param {object} [context]
 * @param {Map<string, object>} [context.relationships] - byId map of the document rels
 * @returns {object}
 */
export function describeSection(sectPrXml, context = {}) {
  const inner = extractElements(sectPrXml, 'w:sectPr')[0]?.innerXml ?? sectPrXml
  const pgSz = extractElements(inner, 'w:pgSz')[0]?.outerXml ?? ''
  const pgMar = extractElements(inner, 'w:pgMar')[0]?.outerXml ?? ''
  const cols = extractElements(inner, 'w:cols')[0]?.outerXml ?? ''
  const pgNumType = extractElements(inner, 'w:pgNumType')[0]?.outerXml ?? ''

  const width = numberOr(getAttribute(pgSz, 'w:w'), null)
  const height = numberOr(getAttribute(pgSz, 'w:h'), null)
  const assumed = width === null || height === null
  const effectiveWidth = width ?? A4_TWIPS.width
  const effectiveHeight = height ?? A4_TWIPS.height

  const margin = (name) => {
    const raw = getAttribute(pgMar, `w:${name}`)
    return { twips: numberOr(raw, null), cm: twipsToCm(raw) }
  }

  const headerRefs = {}
  const footerRefs = {}
  for (const ref of extractElements(inner, 'w:headerReference')) {
    const type = (getAttribute(ref.outerXml, 'w:type') || 'default').toLowerCase()
    const id = getAttribute(ref.outerXml, 'r:id')
    headerRefs[type] = { relId: id, partName: resolveRel(context.relationships, id) }
  }
  for (const ref of extractElements(inner, 'w:footerReference')) {
    const type = (getAttribute(ref.outerXml, 'w:type') || 'default').toLowerCase()
    const id = getAttribute(ref.outerXml, 'r:id')
    footerRefs[type] = { relId: id, partName: resolveRel(context.relationships, id) }
  }

  return {
    type: getAttribute(extractElements(inner, 'w:type')[0]?.outerXml ?? '', 'w:val') || null,
    pageSize: {
      widthTwips: width,
      heightTwips: height,
      widthCm: twipsToCm(effectiveWidth),
      heightCm: twipsToCm(effectiveHeight),
      orientation: effectiveWidth > effectiveHeight ? 'landscape' : 'portrait',
      code: getAttribute(pgSz, 'w:code') === null ? null : Number(getAttribute(pgSz, 'w:code')),
      assumed
    },
    margins: {
      top: margin('top'),
      right: margin('right'),
      bottom: margin('bottom'),
      left: margin('left'),
      header: margin('header'),
      footer: margin('footer'),
      gutter: margin('gutter')
    },
    columns: cols
      ? {
          count: numberOr(getAttribute(cols, 'w:num'), 1),
          spaceTwips: numberOr(getAttribute(cols, 'w:space'), null),
          spaceCm: twipsToCm(getAttribute(cols, 'w:space')),
          separator: /w:sep="(?:1|true)"/i.test(cols)
        }
      : null,
    titlePg: /<w:titlePg(?=[\s/>])/i.test(inner),
    headers: headerRefs,
    footers: footerRefs,
    pageNumbering: pgNumType
      ? {
          start: numberOr(getAttribute(pgNumType, 'w:start'), null),
          format: getAttribute(pgNumType, 'w:fmt') || null
        }
      : null,
    verticalAlign: getAttribute(extractElements(inner, 'w:vAlign')[0]?.outerXml ?? '', 'w:val') || null,
    docGrid: getAttribute(extractElements(inner, 'w:docGrid')[0]?.outerXml ?? '', 'w:type') || null
  }
}

/** Resolve a relationship id to a part name. */
function resolveRel(relationships, id) {
  if (!relationships || !id) return null
  const rel = relationships.get ? relationships.get(id) : relationships[id]
  if (!rel) return null
  const target = rel.target || ''
  if (target.startsWith('/')) return target.replace(/^\/+/, '')
  return target.startsWith('word/') ? target : `word/${target}`
}

/** Numeric attribute with a fallback. */
function numberOr(value, fallback) {
  if (value === null || value === undefined || value === '') return fallback
  const n = Number(value)
  return Number.isFinite(n) ? n : fallback
}

/**
 * Remove every child element with the given tag, optionally only those whose
 * attribute matches. Thin wrapper over the shared ordered-child helper so the
 * section order below is the only thing this module owns.
 * @param {string} inner
 * @param {string} tag
 * @param {{name: string, value: string}} [matchAttr]
 * @returns {string}
 */
export function removeChildren(inner, tag, matchAttr = null) {
  return removeChildElements(inner, tag, matchAttr)
}

/**
 * Insert an element at its schema position inside `<w:sectPr>`.
 * @param {string} inner
 * @param {string} tag
 * @param {string} xml
 * @returns {string}
 */
export function insertChildInOrder(inner, tag, xml) {
  return insertChildOrdered(inner, tag, xml, SECTPR_ORDER)
}

/**
 * Replace (or remove, when `xml` is null) a single-valued child element.
 * @param {string} inner
 * @param {string} tag
 * @param {string|null} xml
 * @returns {string}
 */
export function upsertChild(inner, tag, xml) {
  return upsertChildOrdered(inner, tag, xml, SECTPR_ORDER)
}

/**
 * Apply normalized page-setup changes to an existing `<w:sectPr>` (or build a
 * new one), preserving every property the caller did not name.
 *
 * @param {string|null} sectPrXml - existing element, or null to create one
 * @param {object} spec
 * @param {string} [spec.orientation] - `portrait` | `landscape`
 * @param {number} [spec.widthCm]
 * @param {number} [spec.heightCm]
 * @param {number} [spec.widthTwips]
 * @param {number} [spec.heightTwips]
 * @param {object} [spec.margins] - cm values: `{ top, right, bottom, left, header, footer, gutter }`
 * @param {string} [spec.type] - section break type
 * @param {number} [spec.columns]
 * @param {number} [spec.columnSpaceCm]
 * @param {boolean} [spec.separator]
 * @param {boolean} [spec.titlePg]
 * @param {number} [spec.pageNumberStart]
 * @param {string} [spec.pageNumberFormat]
 * @param {string} [spec.verticalAlign]
 * @returns {string} the full `<w:sectPr>` element
 */
export function applySectionSettings(sectPrXml, spec = {}) {
  let inner = sectPrXml ? (extractElements(sectPrXml, 'w:sectPr')[0]?.innerXml ?? '') : ''

  if (spec.type) {
    const type = String(spec.type)
    if (!SECTION_TYPES.has(type)) {
      throw new Error(`Invalid section type "${spec.type}". Use one of: ${[...SECTION_TYPES].join(', ')}`)
    }
    inner = upsertChild(inner, 'w:type', `<w:type w:val="${type}"/>`)
  }

  const wantsPageSize = spec.orientation !== undefined || spec.widthCm !== undefined
    || spec.heightCm !== undefined || spec.widthTwips !== undefined || spec.heightTwips !== undefined

  if (wantsPageSize) {
    const current = extractElements(inner, 'w:pgSz')[0]
    let width = numberOr(getAttribute(current?.outerXml ?? '', 'w:w'), A4_TWIPS.width)
    let height = numberOr(getAttribute(current?.outerXml ?? '', 'w:h'), A4_TWIPS.height)
    const currentCode = getAttribute(current?.outerXml ?? '', 'w:code')

    if (spec.widthTwips !== undefined && spec.widthTwips !== null) width = Math.round(Number(spec.widthTwips))
    if (spec.heightTwips !== undefined && spec.heightTwips !== null) height = Math.round(Number(spec.heightTwips))
    if (spec.widthCm !== undefined && spec.widthCm !== null) width = cmToTwips(spec.widthCm)
    if (spec.heightCm !== undefined && spec.heightCm !== null) height = cmToTwips(spec.heightCm)

    if (spec.orientation) {
      const orientation = String(spec.orientation).toLowerCase()
      if (orientation !== 'portrait' && orientation !== 'landscape') {
        throw new Error(`Invalid orientation "${spec.orientation}". Use portrait or landscape.`)
      }
      const currentlyLandscape = width > height
      if (orientation === 'landscape' && !currentlyLandscape) [width, height] = [height, width]
      if (orientation === 'portrait' && currentlyLandscape) [width, height] = [height, width]
    }

    const code = currentCode ? ` w:code="${currentCode}"` : ''
    inner = upsertChild(inner, 'w:pgSz', `<w:pgSz w:w="${width}" w:h="${height}"${code}/>`)
  }

  if (spec.margins && typeof spec.margins === 'object') {
    const current = extractElements(inner, 'w:pgMar')[0]
    const read = (name) => numberOr(getAttribute(current?.outerXml ?? '', `w:${name}`), null)
    const values = {
      top: read('top'),
      right: read('right'),
      bottom: read('bottom'),
      left: read('left'),
      header: read('header'),
      footer: read('footer'),
      gutter: read('gutter')
    }
    for (const [name, value] of Object.entries(spec.margins)) {
      if (value === undefined || value === null) continue
      if (!(name in values)) {
        throw new Error(`Unknown margin "${name}". Use one of: ${Object.keys(values).join(', ')}`)
      }
      values[name] = typeof value === 'object'
        ? numberOr(value.twips ?? pointsToTwips(value.points ?? value.pt), null)
        : cmToTwips(value)
    }
    const serialized = Object.entries(values)
      .filter(([, value]) => value !== null)
      .map(([name, value]) => ` w:${name}="${Math.round(value)}"`)
      .join('')
    if (serialized) inner = upsertChild(inner, 'w:pgMar', `<w:pgMar${serialized}/>`)
  }

  if (spec.columns !== undefined || spec.columnSpaceCm !== undefined || spec.separator !== undefined) {
    const current = extractElements(inner, 'w:cols')[0]
    const count = spec.columns ?? numberOr(getAttribute(current?.outerXml ?? '', 'w:num'), 1)
    const space = spec.columnSpaceCm !== undefined && spec.columnSpaceCm !== null
      ? cmToTwips(spec.columnSpaceCm)
      : numberOr(getAttribute(current?.outerXml ?? '', 'w:space'), 708)
    const separator = spec.separator ?? /w:sep="(?:1|true)"/i.test(current?.outerXml ?? '')
    inner = upsertChild(
      inner,
      'w:cols',
      `<w:cols w:num="${Math.round(count)}" w:space="${Math.round(space)}"${separator ? ' w:sep="1"' : ''}/>`
    )
  }

  if (spec.titlePg !== undefined) {
    inner = upsertChild(inner, 'w:titlePg', spec.titlePg ? '<w:titlePg/>' : null)
  }

  if (spec.pageNumberStart !== undefined || spec.pageNumberFormat !== undefined) {
    const current = extractElements(inner, 'w:pgNumType')[0]
    const start = spec.pageNumberStart ?? numberOr(getAttribute(current?.outerXml ?? '', 'w:start'), null)
    const format = spec.pageNumberFormat ?? getAttribute(current?.outerXml ?? '', 'w:fmt')
    if (format && !PAGE_NUMBER_FORMATS.has(format)) {
      throw new Error(`Invalid page number format "${format}". Use one of: ${[...PAGE_NUMBER_FORMATS].join(', ')}`)
    }
    const parts = []
    if (start !== null && start !== undefined) parts.push(`w:start="${Math.round(start)}"`)
    if (format) parts.push(`w:fmt="${format}"`)
    inner = upsertChild(inner, 'w:pgNumType', parts.length > 0 ? `<w:pgNumType ${parts.join(' ')}/>` : null)
  }

  if (spec.verticalAlign !== undefined) {
    inner = upsertChild(
      inner,
      'w:vAlign',
      spec.verticalAlign ? `<w:vAlign w:val="${spec.verticalAlign}"/>` : null
    )
  }

  return `<w:sectPr>${inner}</w:sectPr>`
}

/**
 * Replace the `<w:sectPr>` element at a given position in the document.
 *
 * @param {string} docXml
 * @param {number} index - section index from {@link collectSections}
 * @param {string} newSectPrXml
 * @returns {string}
 */
export function replaceSection(docXml, index, newSectPrXml) {
  const sections = collectSections(docXml)
  if (index < 0 || index >= sections.length) {
    throw new Error(`Section index out of range: ${index} (document has ${sections.length})`)
  }
  const target = sections[index]
  const at = target.position
  return docXml.slice(0, at) + newSectPrXml + docXml.slice(at + target.outerXml.length)
}

/**
 * Insert a section break after a paragraph, returning the new document XML.
 *
 * The section that ends at the break keeps the document's current properties;
 * the request's page-setup changes are applied to the *following* (final)
 * section, which is the behaviour a caller asking for "a landscape second
 * section" expects.
 *
 * @param {string} docXml
 * @param {object} options
 * @param {number} options.afterParagraphIndex - last paragraph of the closed section
 * @param {string} [options.type='nextPage']
 * @param {object} [options.page] - page-setup changes for the following section
 * @returns {{xml: string, breakParagraphIndex: number, sectionIndex: number}}
 */
export function insertSectionBreak(docXml, options = {}) {
  const paragraphs = extractElements(docXml, 'w:p')
  const { afterParagraphIndex, type = 'nextPage' } = options
  if (!Number.isInteger(afterParagraphIndex) || afterParagraphIndex < 0 || afterParagraphIndex >= paragraphs.length) {
    throw new Error(`Paragraph index out of range for a section break: ${afterParagraphIndex} (total: ${paragraphs.length})`)
  }

  const bodySection = extractElements(docXml, 'w:body')[0]
  const finalSectPr = extractElements(bodySection?.innerXml ?? '', 'w:sectPr')[0]
  if (!finalSectPr) {
    throw new Error('Invalid DOCX: the document body has no final <w:sectPr> to split')
  }

  // The paragraph that ends the first section carries a copy of the current
  // final section's properties, with an explicit break type.
  const closing = applySectionSettings(finalSectPr.outerXml, { type })
  const breakParagraph = `<w:p><w:pPr>${closing}</w:pPr></w:p>`

  const anchor = paragraphs[afterParagraphIndex]
  const anchorEnd = anchor.index + anchor.outerXml.length
  let xml = docXml.slice(0, anchorEnd) + breakParagraph + docXml.slice(anchorEnd)

  // Then the changes requested by the caller land on the final section.
  if (options.page && Object.keys(options.page).length > 0) {
    const updated = applySectionSettings(finalSectPr.outerXml, options.page)
    const sections = collectSections(xml)
    const final = sections[sections.length - 1]
    xml = xml.slice(0, final.position) + updated + xml.slice(final.position + final.outerXml.length)
  }

  return { xml, breakParagraphIndex: afterParagraphIndex + 1, sectionIndex: collectSections(xml).length - 1 }
}

/**
 * Every page break in the document.
 *
 * Reports both an explicit `<w:br w:type="page"/>` and a paragraph carrying
 * `<w:pageBreakBefore/>`, because both start a new page and a caller looking
 * for the document's page boundaries wants them together.
 *
 * @param {string} docXml
 * @returns {Array<{paragraphIndex: number, kind: 'break'|'pageBreakBefore', breaks: number}>}
 */
export function listPageBreaks(docXml) {
  const result = []
  const paragraphs = extractElements(docXml, 'w:p')
  for (let i = 0; i < paragraphs.length; i++) {
    const pXml = paragraphs[i].outerXml
    const explicit = (pXml.match(/<w:br(?=[\s/>])[^>]*w:type="page"/gi) || []).length
    const before = /<w:pageBreakBefore(?=[\s/>])([^>]*)/i.test(pXml)
    if (explicit > 0 || before) {
      result.push({ paragraphIndex: i, kind: explicit > 0 ? 'break' : 'pageBreakBefore', breaks: explicit })
    }
  }
  return result
}

/**
 * A header or footer reference as a `<w:headerReference>` element.
 * @param {'header'|'footer'} kind
 * @param {string} relId
 * @param {'default'|'first'|'even'} type
 * @returns {string}
 */
export function buildPartReference(kind, relId, type = 'default') {
  const tag = kind === 'footer' ? 'w:footerReference' : 'w:headerReference'
  return `<${tag} w:type="${type}" r:id="${relId}"/>`
}

/**
 * Add or replace a header/footer reference inside a `<w:sectPr>`.
 * @param {string} sectPrXml
 * @param {object} spec
 * @param {'header'|'footer'} spec.kind
 * @param {string} spec.relId
 * @param {'default'|'first'|'even'} [spec.type]
 * @returns {string}
 */
export function setPartReference(sectPrXml, spec) {
  const kind = spec.kind === 'footer' ? 'footer' : 'header'
  const tag = kind === 'footer' ? 'w:footerReference' : 'w:headerReference'
  const type = spec.type || 'default'
  const inner = extractElements(sectPrXml, 'w:sectPr')[0]?.innerXml ?? ''
  const cleaned = removeChildren(inner, tag, { name: 'w:type', value: type })
  return `<w:sectPr>${insertChildInOrder(cleaned, tag, buildPartReference(kind, spec.relId, type))}</w:sectPr>`
}

/**
 * Append a block of body XML at the end of the document body — after every
 * paragraph and before the *final* `<w:sectPr>`.
 *
 * Slicing at the last section rather than at the first `<w:sectPr` substring
 * matters as soon as the document has a section break: the first matching
 * element belongs to the break paragraph in the middle of the document, and
 * inserting there would drop new content into the wrong section.
 *
 * @param {string} docXml
 * @param {string} xml
 * @returns {string}
 */
export function appendToBodyEnd(docXml, xml) {
  const sections = collectSections(docXml)
  const final = [...sections].reverse().find(section => section.kind === 'final')
  if (final) {
    return docXml.slice(0, final.position) + `${xml}\n` + docXml.slice(final.position)
  }
  if (docXml.includes('</w:body>')) {
    return docXml.replace('</w:body>', `${xml}\n</w:body>`)
  }
  return docXml
}

/**
 * Plain text of a header/footer part, with page-field placeholders preserved
 * as a stable token so a caller can see where the field sits.
 * @param {string} partXml
 * @returns {string}
 */
export function partText(partXml) {
  if (!partXml) return ''
  return extractTextFromXml(partXml.replace(/<w:fldSimple\b[^>]*w:instr="([^"]*)"[^>]*>[\s\S]*?<\/w:fldSimple>/gi, (m, instr) => `{${instr.trim()}}`))
    .replace(/\n{2,}/g, '\n')
    .trim()
}
