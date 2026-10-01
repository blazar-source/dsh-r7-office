/**
 * DOCX list support: `word/numbering.xml`.
 *
 * A bulleted or numbered list in OOXML is two facts: a paragraph carrying
 * `<w:numPr><w:numId …/>` and a numbering definition in `word/numbering.xml`
 * that the id resolves to. The stock R7 template has no numbering part at all,
 * so authoring a list means creating one (plus its relationship and content
 * type) — and when a document already has one, its existing definitions must
 * be reused rather than shadowed.
 *
 * The module keeps that policy explicit: existing `numId`s whose level 0 is
 * already a bullet (or already decimal) are reused, and only genuinely missing
 * definitions are appended.
 */

import {
  CONTENT_TYPES,
  DOCUMENT_PART,
  DOCUMENT_RELS_PART,
  REL_TYPES,
  addRelationship,
  ensureOverrideContentType,
  relativeTarget
} from './docx-parts.js'
import { parseNumberingDefinitions } from './docx-styles.js'
import { extractElements, extractTextFromXml, getAttribute } from '../shared/xml.js'

/** Package part name of the numbering definitions. */
export const NUMBERING_PART = 'word/numbering.xml'

/** Bullet glyphs per level, matching the conventional Word cycle. */
const BULLET_GLYPHS = ['\u2022', 'o', '\u25AA', '\u2022', 'o', '\u25AA', '\u2022', 'o', '\u25AA']

/** Numbering formats the engine can author. */
const NUMBER_FORMATS = {
  decimal: (level) => `%${level + 1}.`,
  lowerLetter: (level) => `%${level + 1})`,
  upperLetter: (level) => `%${level + 1})`,
  lowerRoman: (level) => `%${level + 1}.`,
  upperRoman: (level) => `%${level + 1}.`
}

/**
 * Build one `<w:lvl>` element.
 * @param {number} ilvl
 * @param {string} format
 * @returns {string}
 */
function buildLevel(ilvl, format) {
  const isBullet = format === 'bullet'
  const glyph = isBullet ? BULLET_GLYPHS[ilvl % BULLET_GLYPHS.length] : NUMBER_FORMATS[format](ilvl)
  const indent = 720 + ilvl * 360
  const font = isBullet
    ? `<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>`
    : ''
  return `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/>`
    + `<w:lvlText w:val="${glyph}"/><w:lvlJc w:val="left"/>`
    + `<w:pPr><w:ind w:left="${indent}" w:hanging="360"/></w:pPr>${font}</w:lvl>`
}

/**
 * A complete, self-contained numbering part with one bullet and one decimal
 * definition, each with nine levels.
 * @returns {string}
 */
export function buildNumberingPartXml() {
  const bulletLevels = Array.from({ length: 9 }, (_, i) => buildLevel(i, 'bullet')).join('')
  const decimalLevels = Array.from({ length: 9 }, (_, i) => buildLevel(i, 'decimal')).join('')
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLevels}</w:abstractNum>`
    + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${decimalLevels}</w:abstractNum>`
    + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
    + '</w:numbering>'
}

/** Highest `abstractNumId` used in a numbering part. */
function maxAbstractNumId(xml) {
  let max = -1
  for (const element of extractElements(xml, 'w:abstractNum')) {
    const id = Number(getAttribute(element.outerXml, 'w:abstractNumId'))
    if (Number.isFinite(id)) max = Math.max(max, id)
  }
  return max
}

/** Highest `numId` used in a numbering part. */
function maxNumId(xml) {
  let max = 0
  for (const element of extractElements(xml, 'w:num')) {
    const id = Number(getAttribute(element.outerXml, 'w:numId'))
    if (Number.isFinite(id)) max = Math.max(max, id)
  }
  return max
}

/**
 * Find an existing `numId` whose level 0 uses the requested format.
 * @param {string} numberingXml
 * @param {string} format
 * @returns {number|null}
 */
export function findNumIdByFormat(numberingXml, format) {
  const definitions = parseNumberingDefinitions(numberingXml)
  for (const [numId, definition] of definitions) {
    if (definition.levels?.[0]?.format === format) return numId
  }
  return null
}

/**
 * Append a new abstract numbering definition and bind it to a fresh numId.
 * @param {string} numberingXml
 * @param {string} format
 * @returns {{xml: string, numId: number}}
 */
export function appendNumberingDefinition(numberingXml, format) {
  const abstractId = maxAbstractNumId(numberingXml) + 1
  const numId = maxNumId(numberingXml) + 1
  const levels = Array.from({ length: 9 }, (_, i) => buildLevel(i, format)).join('')
  const abstract = `<w:abstractNum w:abstractNumId="${abstractId}">`
    + `<w:multiLevelType w:val="hybridMultilevel"/>${levels}</w:abstractNum>`
  const num = `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`

  // Schema order matters: every abstractNum precedes every num.
  let xml
  const firstNum = numberingXml.search(/<w:num(?=[\s>])/)
  if (firstNum !== -1) {
    xml = numberingXml.slice(0, firstNum) + abstract + numberingXml.slice(firstNum)
  } else if (numberingXml.includes('</w:numbering>')) {
    xml = numberingXml.replace('</w:numbering>', `${abstract}</w:numbering>`)
  } else {
    throw new Error('Invalid DOCX: word/numbering.xml has no </w:numbering> close tag')
  }

  if (xml.includes('</w:numbering>')) {
    xml = xml.replace('</w:numbering>', `${num}</w:numbering>`)
  }

  return { xml, numId }
}

/**
 * Ensure the document can express a bullet and a numbered list, mutating the
 * archive in place (part, relationship and content type only when missing).
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {object} [options]
 * @param {boolean} [options.bullet=true]
 * @param {boolean} [options.decimal=true]
 * @returns {{bullet: number|null, decimal: number|null, partName: string, created: boolean, added: boolean}}
 */
export function ensureListNumbering(zip, options = {}) {
  const wantBullet = options.bullet !== false
  const wantDecimal = options.decimal !== false

  let created = false
  let numberingXml = zip.getText(NUMBERING_PART)
  if (!numberingXml) {
    numberingXml = buildNumberingPartXml()
    created = true
  }

  let added = false
  let bulletId = findNumIdByFormat(numberingXml, 'bullet')
  let decimalId = findNumIdByFormat(numberingXml, 'decimal')

  if (wantBullet && bulletId === null) {
    const appended = appendNumberingDefinition(numberingXml, 'bullet')
    numberingXml = appended.xml
    bulletId = appended.numId
    added = true
  }
  if (wantDecimal && decimalId === null) {
    const appended = appendNumberingDefinition(numberingXml, 'decimal')
    numberingXml = appended.xml
    decimalId = appended.numId
    added = true
  }

  if (created || added) {
    zip.setText(NUMBERING_PART, numberingXml)
  }

  if (created) {
    // The document part must point at the numbering part.
    const relsXml = zip.getText(DOCUMENT_RELS_PART) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    const withRelationship = addRelationship(relsXml, {
      type: REL_TYPES.numbering,
      target: relativeTarget(DOCUMENT_PART, NUMBERING_PART)
    })
    zip.setText(DOCUMENT_RELS_PART, withRelationship.xml)

    const contentTypes = zip.getText('[Content_Types].xml')
    if (contentTypes) {
      const ensured = ensureOverrideContentType(contentTypes, `/${NUMBERING_PART}`, CONTENT_TYPES.numbering)
      if (ensured.added) zip.setText('[Content_Types].xml', ensured.xml)
    }
  }

  return { bullet: bulletId, decimal: decimalId, partName: NUMBERING_PART, created, added }
}

/**
 * All list paragraphs of a document, with their marker and text.
 * @param {string} docXml
 * @param {string|null} numberingXml
 * @returns {Array<object>}
 */
export function listLists(docXml, numberingXml) {
  const definitions = parseNumberingDefinitions(numberingXml)
  const result = []
  const paragraphs = extractElements(docXml, 'w:p')
  for (let i = 0; i < paragraphs.length; i++) {
    const pXml = paragraphs[i].outerXml
    const numPr = extractElements(pXml, 'w:numPr')[0]
    if (!numPr) continue
    const numId = Number(getAttribute(extractElements(numPr.innerXml, 'w:numId')[0]?.outerXml ?? '', 'w:val'))
    const ilvl = Number(getAttribute(extractElements(numPr.innerXml, 'w:ilvl')[0]?.outerXml ?? '', 'w:val') || 0)
    const level = definitions.get(numId)?.levels?.[ilvl] || null
    result.push({
      paragraphIndex: i,
      numId,
      level: ilvl,
      format: level?.format || null,
      marker: level?.text || null,
      text: extractTextFromXml(pXml).trim()
    })
  }
  return result
}
