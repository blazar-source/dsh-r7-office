/**
 * DOCX headers and footers: `word/header*.xml`, `word/footer*.xml` and the
 * `<w:headerReference>` / `<w:footerReference>` elements inside `<w:sectPr>`.
 *
 * The single most important rule in this module is that a page-number field is
 * not text. Headers and footers exist to carry fields (`PAGE`, `NUMPAGES`,
 * `DATE`, `STYLEREF`), and rewriting a footer by dropping its `<w:t>` values
 * would silently turn "Page 3 of 12" into "Page of" while still looking like a
 * successful edit. Every text rewrite here therefore works on the paragraph's
 * **non-field** runs only: the field elements keep their position, their
 * instructions and their cached results.
 *
 * Both spellings of a field are recognized — the `<w:fldSimple>` element form
 * and the complex `<w:fldChar w:fldCharType="begin"/> … "end"` run form — and
 * a part that gains a page number is written in the `<w:fldSimple>` form.
 */

import {
  CONTENT_TYPES,
  DOCUMENT_PART,
  DOCUMENT_RELS_PART,
  REL_TYPES,
  addRelationship,
  childElements,
  elementParts,
  ensureOverrideContentType,
  escapeRegExp,
  expandSelfClosing,
  nextPartName,
  parseRelationships,
  relativeTarget,
  resolveTarget
} from './docx-parts.js'
import { collectSections, replaceSection, setPartReference } from './docx-sections.js'
import { extractElements, getAttribute } from '../shared/xml.js'
import { normalizeAlignment } from './docx-styles.js'

/** Paragraph-ish parts and their root element name. */
const PART_ROOTS = { header: 'w:hdr', footer: 'w:ftr' }

/** The XML declaration every new part starts with. */
const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

/** Namespaces a new header/footer part needs. */
const PART_NAMESPACES = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

/**
 * Build a complete, minimal header/footer part.
 * @param {'header'|'footer'} kind
 * @returns {string}
 */
export function buildPartXml(kind) {
  const root = PART_ROOTS[kind === 'footer' ? 'footer' : 'header']
  return `${XML_DECLARATION}<${root} ${PART_NAMESPACES}><w:p/></${root}>`
}

/**
 * Build a `<w:fldSimple>` field.
 *
 * The instruction is padded with spaces, which is how Word and R7 write it and
 * how multiple instructions are separated inside one `w:instr` value — a
 * document produced here should be indistinguishable from an authored one.
 *
 * @param {string} instruction - e.g. `PAGE`, `NUMPAGES`, `PAGE of NUMPAGES`
 * @returns {string}
 */
export function buildFieldRun(instruction) {
  const instr = String(instruction).trim()
  return `<w:fldSimple w:instr=" ${instr} "><w:r><w:t>1</w:t></w:r></w:fldSimple>`
}

/**
 * Whether a part contains a field whose instruction mentions `field`.
 *
 * Handles both the `<w:fldSimple w:instr=" PAGE ">` attribute form and the
 * complex `<w:instrText> PAGE </w:instrText>` content form.
 *
 * @param {string} partXml
 * @param {string} [field='PAGE']
 * @returns {boolean}
 */
export function hasField(partXml, field = 'PAGE') {
  if (!partXml) return false
  const simple = new RegExp(
    `<w:fldSimple\\b(?=[^>]*w:instr=["'][^"']*\\b${escapeRegExp(field)}\\b)[^>]*>`,
    'i'
  )
  if (simple.test(partXml)) return true
  const complex = new RegExp(
    `<w:instrText\\b[^>]*>\\s*[^<]*\\b${escapeRegExp(field)}\\b[^<]*</w:instrText>`,
    'i'
  )
  return complex.test(partXml)
}

/**
 * Every field in a header/footer part, in document order.
 * @param {string} partXml
 * @returns {Array<{instruction: string, form: 'fldSimple'|'complex'}>}
 */
export function listFields(partXml) {
  if (!partXml) return []
  const fields = []

  for (const element of extractElements(partXml, 'w:fldSimple')) {
    fields.push({
      instruction: (getAttribute(element.outerXml, 'w:instr') || '').trim(),
      form: 'fldSimple'
    })
  }

  // A complex field spans runs: `begin` … `<w:instrText>` … `separate` … `end`.
  const complexFieldRegex = new RegExp(
    '<w:fldChar\\b[^>]*w:fldCharType="begin"[^>]*/>(.*?)<w:fldChar\\b[^>]*w:fldCharType="end"[^>]*/>',
    'gis'
  )
  let match
  while ((match = complexFieldRegex.exec(partXml)) !== null) {
    const instruction = [...match[1].matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi)]
      .map(entry => entry[1])
      .join('')
      .trim()
    if (instruction) fields.push({ instruction, form: 'complex' })
  }

  return fields
}

/**
 * Whether a run element opens or closes a complex field.
 * @param {string} runXml
 * @returns {'begin'|'end'|null}
 */
function fieldCharKind(runXml) {
  const match = /<w:fldChar\b[^>]*w:fldCharType="(begin|end)"/i.exec(runXml)
  return match ? match[1].toLowerCase() : null
}

/**
 * Rewrite a paragraph's non-field text while leaving every field untouched.
 *
 * The new text lands in the first text-bearing, non-field element (a run, or a
 * hyperlink wrapping runs); any other non-field `<w:t>` in the paragraph is
 * emptied, so the paragraph reads exactly as the caller asked without leaving
 * stale fragments beside the new text. When the paragraph carries no non-field
 * text at all, a fresh run is inserted before the first field so a footer like
 * "Page ⟨PAGE⟩" keeps its field last.
 *
 * @param {string} pXml
 * @param {string} newText
 * @returns {string}
 */
export function rewriteParagraphText(pXml, newText) {
  pXml = expandSelfClosing(pXml)
  const elements = childElements(pXml)
  const pPr = elements.find(element => element.tag === 'w:pPr') || null
  const content = elements.filter(element => element.tag !== 'w:pPr')

  const staleSpans = []
  let firstTarget = null
  let insideComplexField = false

  for (const element of content) {
    if (element.tag === 'w:fldSimple') continue

    const kind = fieldCharKind(element.xml)
    if (kind === 'begin') {
      insideComplexField = true
      continue
    }
    if (insideComplexField) {
      // The run that closes the field is part of it and must survive verbatim.
      if (kind === 'end') insideComplexField = false
      continue
    }

    const tElements = extractElements(element.xml, 'w:t')
    if (tElements.length === 0) continue
    if (!firstTarget) {
      firstTarget = { element, tag: tElements[0] }
    } else {
      for (const t of tElements) {
        staleSpans.push({ offset: element.start + t.index, length: t.outerXml.length })
      }
    }
  }

  let result = pXml
  // Rewrite the trailing text from the end so earlier offsets stay valid.
  for (let i = staleSpans.length - 1; i >= 0; i--) {
    const span = staleSpans[i]
    result = result.slice(0, span.offset) + '<w:t></w:t>' + result.slice(span.offset + span.length)
  }

  if (firstTarget) {
    const absolute = firstTarget.element.start + firstTarget.tag.index
    const replacement = `<w:t xml:space="preserve">${escapeXmlText(newText ?? '')}</w:t>`
    return result.slice(0, absolute) + replacement + result.slice(absolute + firstTarget.tag.outerXml.length)
  }

  // No writable text: insert a run at the start of the paragraph content,
  // before any field, so the field keeps trailing position. An empty paragraph
  // (`<w:p/>` or `<w:p></w:p>`) must receive the run *inside* itself.
  const structure = elementParts(result)
  const insertionPoint = pPr
    ? pPr.end
    : (content[0]?.start ?? (structure ? structure.open.length : 0))
  const run = `<w:r><w:t xml:space="preserve">${escapeXmlText(newText ?? '')}</w:t></w:r>`
  return result.slice(0, insertionPoint) + run + result.slice(insertionPoint)
}

/** Escape text content. */
function escapeXmlText(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * Replace the text of one paragraph of a header/footer part.
 *
 * @param {string} partXml
 * @param {string} newText
 * @param {object} [options]
 * @param {number} [options.paragraphIndex=0] - 0-based paragraph inside the part
 * @returns {{xml: string, paragraphIndex: number, paragraphCount: number}}
 */
export function setPartText(partXml, newText, options = {}) {
  const paragraphIndex = options.paragraphIndex ?? 0
  const paragraphs = extractElements(partXml, 'w:p')
  if (paragraphs.length === 0) {
    throw new Error('Cannot set header/footer text: the part has no paragraphs')
  }
  if (paragraphIndex < 0 || paragraphIndex >= paragraphs.length) {
    throw new Error(`Header/footer paragraph index out of range: ${paragraphIndex} (total: ${paragraphs.length})`)
  }

  const target = paragraphs[paragraphIndex]
  const rewritten = rewriteParagraphText(target.outerXml, newText)
  const xml = partXml.slice(0, target.index) + rewritten + partXml.slice(target.index + target.outerXml.length)
  return { xml, paragraphIndex, paragraphCount: paragraphs.length }
}

/**
 * Append a paragraph (optionally carrying a field) to a header/footer part.
 * @param {string} partXml
 * @param {object} spec
 * @param {string} [spec.text]
 * @param {string} [spec.field] - field instruction, e.g. `PAGE`
 * @param {string} [spec.alignment]
 * @returns {string}
 */
export function appendPartParagraph(partXml, spec = {}) {
  const parts = []
  if (spec.text) parts.push(`<w:r><w:t xml:space="preserve">${escapeXmlText(spec.text)}</w:t></w:r>`)
  if (spec.field) parts.push(buildFieldRun(spec.field))

  const alignment = normalizeAlignment(spec.alignment)
  const pPr = alignment ? `<w:pPr><w:jc w:val="${alignment}"/></w:pPr>` : ''
  const paragraph = `<w:p>${pPr}${parts.join('')}</w:p>`

  const root = /<w:(hdr|ftr)\b[^>]*>/i.exec(partXml)
  if (!root) throw new Error('Invalid header/footer part: root element not found')
  const closing = `</w:${root[1]}>`
  if (!partXml.includes(closing)) throw new Error(`Invalid header/footer part: ${closing} not found`)
  return partXml.replace(closing, `${paragraph}${closing}`)
}

/**
 * Ensure a paragraph of the part carries a field, appending one when the part
 * has none.
 *
 * @param {string} partXml
 * @param {object} spec
 * @param {number} [spec.paragraphIndex=0]
 * @param {'start'|'end'} [spec.position='end']
 * @param {string} [spec.instruction='PAGE']
 * @param {string} [spec.separator] - literal text placed before the field
 * @returns {{xml: string, added: boolean}}
 */
export function ensurePageNumberField(partXml, spec = {}) {
  const instruction = spec.instruction || 'PAGE'
  if (hasField(partXml, instruction)) return { xml: partXml, added: false }

  const paragraphs = extractElements(partXml, 'w:p')
  const paragraphIndex = spec.paragraphIndex ?? 0
  if (paragraphs.length === 0 || paragraphIndex < 0 || paragraphIndex >= paragraphs.length) {
    return { xml: appendPartParagraph(partXml, { field: instruction }), added: true }
  }

  const target = paragraphs[paragraphIndex]
  const separator = spec.separator
    ? `<w:r><w:t xml:space="preserve">${escapeXmlText(spec.separator)}</w:t></w:r>`
    : ''
  const field = buildFieldRun(instruction)

  const paragraphXml = expandSelfClosing(target.outerXml)
  const elements = childElements(paragraphXml)
  const pPr = elements.find(element => element.tag === 'w:pPr') || null

  let updatedParagraph
  if ((spec.position || 'end') === 'start') {
    const at = pPr ? pPr.end : (elements[0]?.start ?? elementParts(paragraphXml).open.length)
    updatedParagraph = paragraphXml.slice(0, at) + field + separator + paragraphXml.slice(at)
  } else {
    const close = elementParts(paragraphXml).close
    const end = paragraphXml.lastIndexOf(close)
    updatedParagraph = paragraphXml.slice(0, end) + separator + field + paragraphXml.slice(end)
  }

  // Rewriting a self-closing paragraph changes its length, so the splice uses
  // the original element's span, not the expanded one.
  const xml = target.outerXml === paragraphXml
    ? splice(partXml, target, updatedParagraph)
    : partXml.slice(0, target.index) + updatedParagraph + partXml.slice(target.index + target.outerXml.length)

  return { xml, added: true }
}

/** Replace an element in a document by its recorded span. */
function splice(xml, element, replacement) {
  return xml.slice(0, element.index) + replacement + xml.slice(element.index + element.outerXml.length)
}

/** Extract header/footer references from a `<w:sectPr>` element. */
function describeReferences(sectPrXml) {
  const headers = {}
  const footers = {}
  for (const ref of extractElements(sectPrXml, 'w:headerReference')) {
    headers[(getAttribute(ref.outerXml, 'w:type') || 'default').toLowerCase()] = getAttribute(ref.outerXml, 'r:id')
  }
  for (const ref of extractElements(sectPrXml, 'w:footerReference')) {
    footers[(getAttribute(ref.outerXml, 'w:type') || 'default').toLowerCase()] = getAttribute(ref.outerXml, 'r:id')
  }
  return { headers, footers }
}

/** Resolve a relationships target to a package-absolute part name. */
export function absolutePart(target) {
  if (!target) return null
  if (target.startsWith('/')) return target.replace(/^\/+/, '')
  return target.startsWith('word/') ? target : resolveTarget(DOCUMENT_PART, target)
}

/**
 * Every header/footer part referenced by the document, in section order.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} docXml
 * @returns {Array<object>}
 */
export function listParts(zip, docXml) {
  const relsXml = zip.getText(DOCUMENT_RELS_PART) || ''
  const relationships = parseRelationships(relsXml)
  const sections = collectSections(docXml)
  const result = []
  const byKey = new Map()

  const push = (kind, type, relId, sectionIndex) => {
    const key = `${kind}:${relId}`
    if (byKey.has(key)) {
      const existing = byKey.get(key)
      if (!existing.sectionIndexes.includes(sectionIndex)) existing.sectionIndexes.push(sectionIndex)
      return
    }
    const rel = relId ? relationships.byId.get(relId) : null
    const partName = rel ? absolutePart(rel.target) : null
    const partXml = partName ? zip.getText(partName) : null
    const entry = {
      kind,
      type,
      relId,
      partName,
      exists: Boolean(partXml),
      sectionIndexes: [sectionIndex],
      text: partXml ? partText(partXml) : null,
      fields: partXml ? listFields(partXml) : [],
      hasPageNumberField: partXml ? hasField(partXml) : false,
      paragraphCount: partXml ? extractElements(partXml, 'w:p').length : 0
    }
    byKey.set(key, entry)
    result.push(entry)
  }

  for (const section of sections) {
    const refs = describeReferences(section.outerXml)
    for (const [type, relId] of Object.entries(refs.headers)) push('header', type, relId, section.index)
    for (const [type, relId] of Object.entries(refs.footers)) push('footer', type, relId, section.index)
  }

  return result
}

/**
 * A header/footer part's readable text, with fields shown as `{INSTR}` so a
 * caller can see that the page number is still there.
 * @param {string} partXml
 * @returns {string}
 */
export function partText(partXml) {
  if (!partXml) return ''
  const withTokens = partXml.replace(
    /<w:fldSimple\b[^>]*w:instr="([^"]*)"[^>]*>[\s\S]*?<\/w:fldSimple>/gi,
    (m, instr) => `{${instr.trim()}}`
  )
  return withTokens
    .replace(/<w:br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim()
}

/**
 * Ensure a header/footer part exists for a section and is referenced by it.
 * An already-referenced part is left in place and only its text/field is
 * updated, so an existing header is never recreated (which would lose its
 * images, tables or fields).
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} docXml
 * @param {object} options
 * @param {'header'|'footer'} options.kind
 * @param {'default'|'first'|'even'} [options.type='default']
 * @param {number} [options.sectionIndex=0]
 * @param {string} [options.text]
 * @param {string} [options.field] - a field to add when missing, e.g. `PAGE`
 * @param {string} [options.alignment]
 * @returns {{docXml: string, partName: string, relId: string, created: boolean, kind: string, type: string}}
 */
export function ensurePart(zip, docXml, options) {
  const kind = options.kind === 'footer' ? 'footer' : 'header'
  const type = options.type || 'default'
  const sections = collectSections(docXml)
  const sectionIndex = options.sectionIndex ?? 0
  if (sectionIndex < 0 || sectionIndex >= sections.length) {
    throw new Error(`Section index out of range: ${sectionIndex} (document has ${sections.length})`)
  }

  const section = sections[sectionIndex]
  const refs = describeReferences(section.outerXml)
  const existingRelId = (kind === 'header' ? refs.headers : refs.footers)[type]

  if (existingRelId) {
    const relsXml = zip.getText(DOCUMENT_RELS_PART) || ''
    const rel = parseRelationships(relsXml).byId.get(existingRelId)
    const partName = rel ? absolutePart(rel.target) : null
    if (!partName) {
      throw new Error(`Header/footer relationship "${existingRelId}" has no resolvable target`)
    }
    let xml = zip.getText(partName) || buildPartXml(kind)
    const before = xml
    if (options.text !== undefined && options.text !== null && options.text !== '') {
      xml = setPartText(xml, options.text, {}).xml
    }
    if (options.field) {
      xml = ensurePageNumberField(xml, { instruction: options.field }).xml
    }
    if (xml !== before) zip.setText(partName, xml)
    return { docXml, partName, relId: existingRelId, created: false, kind, type }
  }

  const partName = nextPartName(zip, `word/${kind}`, 'xml')
  let partXml = buildPartXml(kind)
  if (options.text) partXml = setPartText(partXml, options.text, {}).xml
  if (options.field) partXml = ensurePageNumberField(partXml, { instruction: options.field }).xml
  if (options.alignment) partXml = applyAlignment(partXml, options.alignment)
  zip.setText(partName, partXml)

  const relsXml = zip.getText(DOCUMENT_RELS_PART)
  const withRelationship = addRelationship(relsXml, {
    type: kind === 'footer' ? REL_TYPES.footer : REL_TYPES.header,
    target: relativeTarget(DOCUMENT_PART, partName)
  })
  zip.setText(DOCUMENT_RELS_PART, withRelationship.xml)

  const contentTypes = zip.getText('[Content_Types].xml')
  if (contentTypes) {
    const ensured = ensureOverrideContentType(
      contentTypes,
      `/${partName}`,
      kind === 'footer' ? CONTENT_TYPES.footer : CONTENT_TYPES.header
    )
    if (ensured.added) zip.setText('[Content_Types].xml', ensured.xml)
  }

  const updatedSection = setPartReference(section.outerXml, { kind, relId: withRelationship.id, type })
  const docWithRef = replaceSection(docXml, sectionIndex, updatedSection)

  return { docXml: docWithRef, partName, relId: withRelationship.id, created: true, kind, type }
}

/** Apply an alignment to the part's first paragraph. */
function applyAlignment(partXml, alignment) {
  const paragraphs = extractElements(partXml, 'w:p')
  if (paragraphs.length === 0) return partXml
  const target = paragraphs[0]
  const paragraphXml = expandSelfClosing(target.outerXml)
  const jc = `<w:jc w:val="${normalizeAlignment(alignment)}"/>`
  const updated = /<w:pPr(?=[\s/>])/.test(paragraphXml)
    ? paragraphXml.replace(/(<w:pPr(?=[\s/>])[^>]*>)/, `$1${jc}`)
    : paragraphXml.replace(/^(<w:p(?=[\s/>])[^>]*>)/, `$1<w:pPr>${jc}</w:pPr>`)
  return partXml.slice(0, target.index) + updated + partXml.slice(target.index + target.outerXml.length)
}

/**
 * Delete a header/footer part reference and its part, leaving the rest of the
 * section untouched.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} docXml
 * @param {object} options
 * @param {number} [options.sectionIndex=0]
 * @param {'header'|'footer'} options.kind
 * @param {'default'|'first'|'even'} [options.type='default']
 * @returns {{docXml: string, removedPart: string|null, removedRelId: string|null}}
 */
export function removePart(zip, docXml, options) {
  const kind = options.kind === 'footer' ? 'footer' : 'header'
  const type = options.type || 'default'
  const sections = collectSections(docXml)
  const sectionIndex = options.sectionIndex ?? 0
  const section = sections[sectionIndex]
  if (!section) throw new Error(`Section index out of range: ${sectionIndex} (document has ${sections.length})`)

  const refs = describeReferences(section.outerXml)
  const relId = (kind === 'header' ? refs.headers : refs.footers)[type]
  if (!relId) return { docXml, removedPart: null, removedRelId: null }

  const relsXml = zip.getText(DOCUMENT_RELS_PART) || ''
  const relElement = extractElements(relsXml, 'Relationship')
    .find(element => getAttribute(element.outerXml, 'Id') === relId)
  const partName = relElement ? absolutePart(getAttribute(relElement.outerXml, 'Target')) : null

  const tag = kind === 'footer' ? 'w:footerReference' : 'w:headerReference'
  const cleaned = section.outerXml.replace(
    new RegExp(`\\s*<${tag}\\b(?=[^>]*w:type="${escapeRegExp(type)}")[^>]*/>`, 'i'),
    ''
  )
  const docWithRef = replaceSection(docXml, sectionIndex, cleaned)

  if (partName) zip.remove(partName)
  // The relationship is dropped in place so no dangling id is left behind.
  if (relElement) {
    zip.setText(
      DOCUMENT_RELS_PART,
      relsXml.slice(0, relElement.index) + relsXml.slice(relElement.index + relElement.outerXml.length)
    )
  }

  return { docXml: docWithRef, removedPart: partName, removedRelId: relId }
}
