/**
 * DOCX formatting model: `word/styles.xml`, `word/numbering.xml` and the
 * `<w:pPr>` / `<w:rPr>` / `<w:tcPr>` property bags inside `word/document.xml`.
 *
 * Two halves live here and they deliberately share one vocabulary:
 *
 *   - **read** — `extractParagraphFormatting`, `extractTableFormatting` and
 *     friends turn raw OOXML properties into normalized JavaScript objects, so
 *     an agent can reason about an existing document without ever seeing a
 *     `<w:spacing>` element;
 *   - **write** — `buildParagraphXml` / `buildRunProperties` take the same
 *     normalized vocabulary and emit OOXML.
 *
 * Everything is optional and absence-aware: a property that the document does
 * not carry reads back as `null` rather than as a guessed default, because
 * "unset" and "explicitly set to the default" are different facts about a
 * document and only the first one should be reported as unknown.
 *
 * Units follow OOXML: lengths are twentieths of a point (twips), font sizes
 * are half-points, borders are eighths of a point. The normalized surface
 * converts all of those to points (and, for colours, to `#RRGGBB`) and keeps a
 * `*Twips` / `raw` companion only where the exact OOXML number matters.
 */

import { escapeXml, extractTextFromXml, extractElements, getAttribute } from '../shared/xml.js'

/** Twips per point. */
export const TWIPS_PER_POINT = 20

/** Points per centimetre. */
export const POINTS_PER_CM = 72 / 2.54

/** Alignments accepted in `<w:jc w:val="…">`. */
export const ALIGNMENTS = new Set(['left', 'center', 'right', 'both', 'distribute', 'start', 'end', 'justify'])

/** Table cell vertical alignments. */
export const VERTICAL_ALIGNMENTS = new Set(['top', 'center', 'bottom'])

/** The paragraph styles the engine can author itself. */
export const BUILTIN_HEADING_STYLE_IDS = ['Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6']

/** Human-readable units of every normalized number this module returns. */
export const DOCX_STYLE_VOCABULARY = {
  note: 'Every value below is normalized; no raw OOXML is needed to change it.',
  runKeys: ['text', 'bold', 'italic', 'underline', 'strike', 'family', 'size', 'color', 'highlight', 'vertAlign', 'rStyle'],
  paragraphKeys: [
    'style', 'styleName', 'isHeading', 'headingLevel', 'alignment', 'indents',
    'spacing', 'list', 'pageBreakBefore', 'keepNext', 'keepLines', 'shading',
    'borders', 'outlineLevel', 'sectionBreak', 'runs'
  ],
  tableCellKeys: ['width', 'gridSpan', 'vMerge', 'shading', 'borders', 'verticalAlign', 'margins', 'alignment', 'text'],
  units: {
    runSize: 'points (half-points in OOXML)',
    indents: 'points (twips in OOXML)',
    spacing: 'points, except line with lineRule "auto" which is a multiple',
    tableWidths: 'twips, plus cm',
    colour: '#RRGGBB or "auto"'
  },
  listFormats: ['bullet', 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'none']
}

/** A colour as `#RRGGBB`, or `auto`, or null when the property is absent. */
export function normalizeColor(value) {
  if (!value || typeof value !== 'string') return null
  const trimmed = value.trim()
  if (/^auto$/i.test(trimmed)) return 'auto'
  const hex = trimmed.replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{6}$/.test(hex)) return `#${hex}`
  if (/^[0-9A-F]{8}$/.test(hex)) return `#${hex.slice(2)}`
  return null
}

/** Convert a `<w:color>` value to `#RRGGBB` / `auto`, or null. */
export function colorToHex(value) {
  return normalizeColor(value)
}

/** Parse a numeric attribute value, tolerating whitespace; null when absent. */
export function toNumber(value) {
  if (value === null || value === undefined || value === '') return null
  const n = Number(String(value).trim())
  return Number.isFinite(n) ? n : null
}

/** Half-points (OOXML `w:sz`) to points. */
export function halfPointsToPoints(value) {
  const n = toNumber(value)
  return n === null ? null : n / 2
}

/** Points to half-points, rounded to the nearest half point. */
export function pointsToHalfPoints(value) {
  const n = toNumber(value)
  return n === null ? null : Math.round(n * 2)
}

/** Twips to points. */
export function twipsToPoints(value) {
  const n = toNumber(value)
  return n === null ? null : n / TWIPS_PER_POINT
}

/** Points to whole twips. */
export function pointsToTwips(value) {
  const n = toNumber(value)
  return n === null ? null : Math.round(n * TWIPS_PER_POINT)
}

/** Centimetres to twips. */
export function cmToTwips(value) {
  const n = toNumber(value)
  return n === null ? null : Math.round(n * POINTS_PER_CM * TWIPS_PER_POINT)
}

/** Twips to centimetres, rounded to two decimals. */
export function twipsToCm(value) {
  const n = toNumber(value)
  return n === null ? null : Math.round((n / TWIPS_PER_POINT / POINTS_PER_CM) * 100) / 100
}

/**
 * Read a toggle property such as `<w:b/>` or `<w:b w:val="0"/>`.
 * @param {string} propsXml - a `<w:rPr>` or `<w:pPr>` fragment
 * @param {string} tag - e.g. `w:b`
 * @returns {boolean|null} null when the property is absent
 */
export function readToggle(propsXml, tag) {
  const regex = new RegExp(`<${tag}(?=[\\s/>])([^>]*)/?>`, 'i')
  const match = regex.exec(propsXml)
  if (!match) return null
  const value = getAttribute(match[1], 'w:val')
  if (value === null) return true
  return !/^(0|false|off|none)$/i.test(value)
}

/**
 * Read a simple `<w:tag w:val="…"/>` property.
 * @param {string} propsXml
 * @param {string} tag
 * @returns {string|null}
 */
export function readValue(propsXml, tag) {
  const regex = new RegExp(`<${tag}(?=[\\s/>])([^>]*?)/?>`, 'i')
  const match = regex.exec(propsXml)
  if (!match) return null
  return getAttribute(match[1], 'w:val')
}

/**
 * Extract the property bag that is a direct child of a paragraph element.
 * @param {string} pXml
 * @param {string} tag - e.g. `w:pPr`
 * @returns {string} the inner XML, or '' when absent
 */
export function propertyBag(pXml, tag) {
  const elements = extractElements(pXml, tag)
  if (elements.length === 0) return ''
  return elements[0].innerXml
}

/**
 * Extract normalized character formatting from a run (or a `<w:rPr>` bag).
 * @param {string} xml - a `<w:r>` element or a bare `<w:rPr>` bag
 * @returns {object}
 */
export function extractRunFormatting(xml) {
  const rPr = xml.includes('<w:rPr') ? (extractElements(xml, 'w:rPr')[0]?.innerXml ?? propertyBag(xml, 'w:rPr')) : ''
  const fonts = extractElements(rPr, 'w:rFonts')[0]
  const fontsTag = fonts ? fonts.outerXml : ''
  // `<w:u/>` without a value means "single"; `<w:u w:val="none"/>` means off.
  const underlineTag = extractElements(rPr, 'w:u')[0]?.outerXml ?? null
  const underline = underlineTag === null ? null : (getAttribute(underlineTag, 'w:val') || 'single')
  const vertAlign = readValue(rPr, 'w:vertAlign')

  return {
    bold: readToggle(rPr, 'w:b') === true,
    italic: readToggle(rPr, 'w:i') === true,
    underline: Boolean(underline && !/^none$/i.test(underline)),
    underlineStyle: underline && !/^none$/i.test(underline) ? underline : null,
    strike: readToggle(rPr, 'w:strike') === true,
    family: fonts
      ? (getAttribute(fontsTag, 'w:ascii') || getAttribute(fontsTag, 'w:hAnsi') || getAttribute(fontsTag, 'w:cs') || null)
      : null,
    size: halfPointsToPoints(getAttribute(extractElements(rPr, 'w:sz')[0]?.outerXml ?? '', 'w:val')),
    color: normalizeColor(getAttribute(extractElements(rPr, 'w:color')[0]?.outerXml ?? '', 'w:val')),
    highlight: getAttribute(extractElements(rPr, 'w:highlight')[0]?.outerXml ?? '', 'w:val'),
    vertAlign: vertAlign && !/^baseline$/i.test(vertAlign) ? vertAlign : null,
    rStyle: readValue(rPr, 'w:rStyle')
  }
}

/**
 * Extract the plain text of a run, including tabs and soft breaks.
 * @param {string} rXml
 * @returns {string}
 */
export function runText(rXml) {
  return extractTextFromXml(rXml)
}

/**
 * Normalized formatting of a single `<w:p>` element.
 *
 * @param {string} pXml
 * @param {object} [context]
 * @param {Map<number, object>} [context.numbering] - numId to parsed numbering definition
 * @param {Map<string, object>} [context.styles] - styleId to parsed style definition
 * @returns {object}
 */
export function extractParagraphFormatting(pXml, context = {}) {
  const pPr = propertyBag(pXml, 'w:pPr')
  const style = readValue(pPr, 'w:pStyle') || 'Normal'
  const styleInfo = context.styles?.get(style) || null

  const numPr = extractElements(pPr, 'w:numPr')[0]?.innerXml ?? ''
  const numId = toNumber(readValue(numPr, 'w:numId'))
  const ilvl = toNumber(readValue(numPr, 'w:ilvl')) ?? 0
  const numbering = numId !== null ? context.numbering?.get(numId) || null : null
  const level = numbering?.levels?.[ilvl] || null

  const ind = extractElements(pPr, 'w:ind')[0]?.outerXml ?? ''
  const spacingTag = extractElements(pPr, 'w:spacing')[0]?.outerXml ?? ''
  const lineRule = getAttribute(spacingTag, 'w:lineRule')
  const rawLine = toNumber(getAttribute(spacingTag, 'w:line'))

  let line = null
  if (rawLine !== null) {
    line = !lineRule || lineRule === 'auto'
      ? Math.round((rawLine / 240) * 100) / 100
      : twipsToPoints(rawLine)
  }

  const indent = (name) => {
    const raw = getAttribute(ind, `w:${name}`)
    return { points: twipsToPoints(raw), twips: toNumber(raw) }
  }
  const spacing = (name) => {
    const raw = getAttribute(spacingTag, `w:${name}`)
    return { points: twipsToPoints(raw), twips: toNumber(raw) }
  }

  const shd = extractElements(pPr, 'w:shd')[0]?.outerXml ?? ''
  const borders = extractBorders(extractElements(pPr, 'w:pBdr')[0]?.innerXml ?? '')
  const styleName = styleInfo?.name || style
  const isHeading = /^heading\s*([1-9])/i.test(styleName) || /^заголовок\s*([1-9])/i.test(styleName)
    || /^heading[1-9]$/i.test(style)

  const runs = extractElements(pXml, 'w:r').map(r => ({
    text: runText(r.outerXml),
    ...extractRunFormatting(r.outerXml)
  }))

  return {
    style,
    styleName,
    isHeading,
    headingLevel: isHeading ? Number(/([1-9])/.exec(styleName)?.[1] || /([1-9])$/.exec(style)?.[1] || 1) : null,
    alignment: normalizeAlignment(readValue(pPr, 'w:jc')),
    indents: {
      left: indent('left').points,
      right: indent('right').points,
      firstLine: indent('firstLine').points,
      hanging: indent('hanging').points
    },
    spacing: {
      before: spacing('before').points,
      after: spacing('after').points,
      line,
      lineRule: lineRule || (rawLine === null ? null : 'auto')
    },
    list: numId === null
      ? null
      : {
          numId,
          level: ilvl,
          format: level?.format || null,
          marker: level?.text || null,
          start: level?.start ?? null
        },
    pageBreakBefore: readToggle(pPr, 'w:pageBreakBefore') === true,
    keepNext: readToggle(pPr, 'w:keepNext') === true,
    keepLines: readToggle(pPr, 'w:keepLines') === true,
    shading: normalizeColor(getAttribute(shd, 'w:fill')),
    borders,
    outlineLevel: toNumber(readValue(pPr, 'w:outlineLvl')),
    sectionBreak: extractElements(pPr, 'w:sectPr').length > 0,
    runs
  }
}

/** Normalize a `<w:jc>` value to a stable alignment name. */
export function normalizeAlignment(value) {
  if (!value) return null
  const v = String(value).toLowerCase()
  if (v === 'start') return 'left'
  if (v === 'end') return 'right'
  if (v === 'justify') return 'both'
  return ALIGNMENTS.has(v) ? v : v
}

/**
 * Extract normalized borders from a `<w:tblBorders>` / `<w:tcBorders>` /
 * `<w:pBdr>` inner XML. Only edges actually present are returned.
 *
 * @param {string} bordersXml
 * @returns {object|null}
 */
export function extractBorders(bordersXml) {
  if (!bordersXml) return null
  const edges = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'start', 'end']
  const result = {}
  for (const edge of edges) {
    const tag = extractElements(bordersXml, `w:${edge}`)[0]?.outerXml
    if (!tag) continue
    const style = getAttribute(tag, 'w:val') || 'single'
    const size = toNumber(getAttribute(tag, 'w:sz'))
    result[edge] = {
      style: /^nil$/i.test(style) ? 'none' : style,
      sizeEighthsOfPoint: size,
      widthPoints: size === null ? null : Math.round((size / 8) * 100) / 100,
      color: normalizeColor(getAttribute(tag, 'w:color')),
      spacePoints: twipsToPoints(getAttribute(tag, 'w:space'))
    }
  }
  return Object.keys(result).length > 0 ? result : null
}

/**
 * Normalized formatting of a `<w:tc>` element.
 * @param {string} tcXml
 * @param {object} [context]
 * @returns {object}
 */
export function extractCellFormatting(tcXml, context = {}) {
  const tcPr = propertyBag(tcXml, 'w:tcPr')
  const tcW = extractElements(tcPr, 'w:tcW')[0]?.outerXml ?? ''
  const shd = extractElements(tcPr, 'w:shd')[0]?.outerXml ?? ''
  const vMerge = extractElements(tcPr, 'w:vMerge')[0]?.outerXml ?? ''
  const vAlign = readValue(tcPr, 'w:vAlign')
  const paragraphs = extractElements(tcXml, 'w:p')
  const first = paragraphs[0]?.outerXml ?? ''
  const firstFormat = first ? extractParagraphFormatting(first, context) : null

  const marginsXml = extractElements(tcPr, 'w:tcMar')[0]?.innerXml ?? ''
  const margin = (name) => twipsToPoints(getAttribute(extractElements(marginsXml, `w:${name}`)[0]?.outerXml ?? '', 'w:w'))

  const widthTwips = toNumber(getAttribute(tcW, 'w:w'))

  return {
    width: tcW
      ? {
          twips: widthTwips,
          points: twipsToPoints(widthTwips),
          cm: twipsToCm(widthTwips),
          mode: getAttribute(tcW, 'w:type') || 'dxa'
        }
      : null,
    gridSpan: toNumber(readValue(tcPr, 'w:gridSpan')) ?? 1,
    vMerge: vMerge
      ? ((getAttribute(vMerge, 'w:val') || 'continue').toLowerCase() === 'restart' ? 'restart' : 'continue')
      : null,
    shading: normalizeColor(getAttribute(shd, 'w:fill')),
    borders: extractBorders(extractElements(tcPr, 'w:tcBorders')[0]?.innerXml ?? ''),
    verticalAlign: vAlign ? vAlign.toLowerCase() : null,
    margins: marginsXml
      ? { top: margin('top'), left: margin('left'), bottom: margin('bottom'), right: margin('right') }
      : null,
    alignment: firstFormat ? firstFormat.alignment : null,
    text: extractTextFromXml(tcXml).trim()
  }
}

/**
 * Normalized description of a `<w:tbl>` element.
 * @param {string} tblXml
 * @param {object} [context]
 * @returns {object}
 */
export function extractTableFormatting(tblXml, context = {}) {
  const tblPr = propertyBag(tblXml, 'w:tblPr')
  const tblW = extractElements(tblPr, 'w:tblW')[0]?.outerXml ?? ''
  const grid = extractElements(tblXml, 'w:tblGrid')[0]?.innerXml ?? ''
  const columns = extractElements(grid, 'w:gridCol').map(c => toNumber(getAttribute(c.outerXml, 'w:w')))

  const rows = extractElements(tblXml, 'w:tr').map(r => {
    const trPr = propertyBag(r.outerXml, 'w:trPr')
    const trHeight = extractElements(trPr, 'w:trHeight')[0]?.outerXml ?? ''
    const cells = extractElements(r.outerXml, 'w:tc').map(c => extractCellFormatting(c.outerXml, context))
    const heightTwips = toNumber(getAttribute(trHeight, 'w:val'))
    return {
      height: trHeight
        ? { twips: heightTwips, points: twipsToPoints(heightTwips), rule: getAttribute(trHeight, 'w:hRule') || 'atLeast' }
        : null,
      header: readToggle(trPr, 'w:tblHeader') === true,
      cantSplit: readToggle(trPr, 'w:cantSplit') === true,
      cells
    }
  })

  const widthTwips = toNumber(getAttribute(tblW, 'w:w'))

  return {
    style: readValue(tblPr, 'w:tblStyle'),
    alignment: normalizeAlignment(readValue(tblPr, 'w:jc')),
    // `w:tblLayout` names its value `w:type`, not `w:val`.
    layout: getAttribute(extractElements(tblPr, 'w:tblLayout')[0]?.outerXml ?? '', 'w:type') || null,
    width: tblW
      ? { twips: widthTwips, cm: twipsToCm(widthTwips), mode: getAttribute(tblW, 'w:type') || 'auto' }
      : null,
    borders: extractBorders(extractElements(tblPr, 'w:tblBorders')[0]?.innerXml ?? ''),
    columns,
    columnWidthsTwips: columns,
    columnCount: columns.length || (rows[0]?.cells.length ?? 0),
    rowCount: rows.length,
    rows
  }
}

/**
 * Parse `word/numbering.xml` into a numId to definition map.
 *
 * @param {string|null} numberingXml
 * @returns {Map<number, {numId: number, abstractNumId: number|null, levels: Array<object>}>}
 */
export function parseNumberingDefinitions(numberingXml) {
  const result = new Map()
  if (!numberingXml) return result

  const abstractById = new Map()
  for (const abstract of extractElements(numberingXml, 'w:abstractNum')) {
    const id = toNumber(getAttribute(abstract.outerXml, 'w:abstractNumId'))
    const levels = []
    for (const lvl of extractElements(abstract.outerXml, 'w:lvl')) {
      const ilvl = toNumber(getAttribute(lvl.outerXml, 'w:ilvl')) ?? levels.length
      const numFmt = readValue(lvl.outerXml, 'w:numFmt')
      const lvlText = readValue(lvl.outerXml, 'w:lvlText')
      const start = toNumber(readValue(lvl.outerXml, 'w:start'))
      levels[ilvl] = { level: ilvl, format: numFmt, text: lvlText, start }
    }
    abstractById.set(id, levels)
  }

  for (const num of extractElements(numberingXml, 'w:num')) {
    const numId = toNumber(getAttribute(num.outerXml, 'w:numId'))
    if (numId === null) continue
    const abstractId = toNumber(readValue(num.innerXml, 'w:abstractNumId'))
    const abstractLevels = abstractId === null ? [] : abstractById.get(abstractId) || []
    const overrides = new Map()
    for (const override of extractElements(num.innerXml, 'w:lvlOverride')) {
      const ilvl = toNumber(getAttribute(override.outerXml, 'w:ilvl'))
      const startOverride = toNumber(readValue(override.innerXml, 'w:startOverride'))
      if (ilvl !== null) overrides.set(ilvl, startOverride)
    }
    const levels = abstractLevels.map((level, index) => (
      overrides.has(index) && level
        ? { ...level, start: overrides.get(index) ?? level.start }
        : level
    ))
    result.set(numId, { numId, abstractNumId: abstractId, levels })
  }

  return result
}

/**
 * Parse `word/styles.xml` into a styleId to definition map.
 * @param {string|null} stylesXml
 * @returns {Map<string, {styleId: string, name: string, type: string, basedOn: string|null, headingLevel: number|null}>}
 */
export function parseStyleDefinitions(stylesXml) {
  const result = new Map()
  if (!stylesXml) return result

  for (const style of extractElements(stylesXml, 'w:style')) {
    const styleId = getAttribute(style.outerXml, 'w:styleId')
    if (!styleId) continue
    const name = readValue(style.innerXml, 'w:name') || styleId
    const headingMatch = /^(?:heading|заголовок)\s*([1-9])$/i.exec(name)
    result.set(styleId, {
      styleId,
      name,
      type: getAttribute(style.outerXml, 'w:type') || 'paragraph',
      basedOn: readValue(style.innerXml, 'w:basedOn'),
      headingLevel: headingMatch ? Number(headingMatch[1]) : null
    })
  }

  return result
}

/**
 * Resolve a friendly style request (`Heading2`, `heading 2`, `Заголовок 2`)
 * against the styles that actually exist in the document, falling back to the
 * canonical Word style id.
 *
 * @param {Map<string, object>} styles - from {@link parseStyleDefinitions}
 * @param {string} requested
 * @returns {string}
 */
export function resolveStyleId(styles, requested) {
  if (!requested) return 'Normal'
  const wanted = String(requested).trim()
  if (wanted.toLowerCase() === 'normal') return styles.has('Normal') ? 'Normal' : wanted
  if (styles?.has(wanted)) return wanted

  const headingMatch = /^(?:heading|заголовок)\s*([1-9])$/i.exec(wanted)
  if (headingMatch) {
    const canonical = `Heading${headingMatch[1]}`
    if (styles?.has(canonical)) return canonical
    for (const [id, def] of styles || []) {
      if (def.headingLevel === Number(headingMatch[1])) return id
    }
    return canonical
  }

  // Match by display name as a last resort, case-insensitively.
  for (const [id, def] of styles || []) {
    if (String(def.name).toLowerCase() === wanted.toLowerCase()) return id
  }
  return wanted
}

/* ------------------------------------------------------------------ *
 * Writers
 * ------------------------------------------------------------------ */

/**
 * Build a `<w:rPr>` fragment from normalized run formatting.
 * Unknown keys are ignored; nothing is emitted for absent properties.
 *
 * @param {object} spec
 * @returns {string} inner XML of `<w:rPr>` (empty string when nothing is set)
 */
export function buildRunProperties(spec = {}) {
  const parts = []
  if (spec.bold) parts.push('<w:b/>')
  if (spec.bold === false) parts.push('<w:b w:val="0"/>')
  if (spec.italic) parts.push('<w:i/>')
  if (spec.italic === false) parts.push('<w:i w:val="0"/>')
  if (spec.strike) parts.push('<w:strike/>')

  const underline = spec.underline === true ? 'single' : (typeof spec.underline === 'string' ? spec.underline : null)
  if (underline) parts.push(`<w:u w:val="${underline}"/>`)

  const family = spec.family || spec.font
  if (family) {
    const escaped = escapeXml(family)
    parts.push(`<w:rFonts w:ascii="${escaped}" w:hAnsi="${escaped}" w:cs="${escaped}"/>`)
  }

  if (spec.size !== undefined && spec.size !== null) {
    const half = pointsToHalfPoints(spec.size)
    if (half !== null) parts.push(`<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`)
  }

  const color = normalizeColor(spec.color)
  if (color) parts.push(`<w:color w:val="${escapeXml(color.replace(/^#/, ''))}"/>`)

  if (spec.highlight) parts.push(`<w:highlight w:val="${escapeXml(spec.highlight)}"/>`)
  if (spec.vertAlign) parts.push(`<w:vertAlign w:val="${escapeXml(spec.vertAlign)}"/>`)
  if (spec.rStyle) parts.push(`<w:rStyle w:val="${escapeXml(spec.rStyle)}"/>`)

  return parts.join('')
}

/**
 * Build a `<w:pPr>` fragment from normalized paragraph formatting.
 *
 * @param {object} spec
 * @param {object} [context]
 * @param {string} [context.styleId] - resolved style id to apply
 * @returns {string} inner XML of `<w:pPr>`
 */
export function buildParagraphProperties(spec = {}, context = {}) {
  const parts = []
  const styleId = context.styleId || spec.styleId || spec.style
  if (styleId && styleId !== 'Normal') parts.push(`<w:pStyle w:val="${escapeXml(styleId)}"/>`)

  if (spec.keepNext) parts.push('<w:keepNext/>')
  if (spec.keepLines) parts.push('<w:keepLines/>')
  if (spec.pageBreakBefore) parts.push('<w:pageBreakBefore/>')

  if (spec.list && spec.list.numId !== undefined && spec.list.numId !== null) {
    const level = spec.list.level ?? spec.list.ilvl ?? 0
    parts.push(`<w:numPr><w:ilvl w:val="${Number(level)}"/><w:numId w:val="${Number(spec.list.numId)}"/></w:numPr>`)
  }

  if (spec.borders && typeof spec.borders === 'object') {
    const edges = ['top', 'left', 'bottom', 'right']
    const xml = edges
      .filter(edge => spec.borders[edge])
      .map(edge => buildBorderEdge(edge, spec.borders[edge]))
      .join('')
    if (xml) parts.push(`<w:pBdr>${xml}</w:pBdr>`)
  }

  const shading = normalizeColor(spec.shading)
  if (shading && shading !== 'auto') parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${shading.replace(/^#/, '')}"/>`)

  const spacingParts = []
  const before = spec.spacing?.before ?? spec.spaceBefore
  const after = spec.spacing?.after ?? spec.spaceAfter
  const beforeLines = spec.spacing?.beforeLines
  const afterLines = spec.spacing?.afterLines
  if (before !== undefined && before !== null) spacingParts.push(`w:before="${pointsToTwips(before)}"`)
  if (after !== undefined && after !== null) spacingParts.push(`w:after="${pointsToTwips(after)}"`)
  if (beforeLines !== undefined && beforeLines !== null) spacingParts.push(`w:beforeLines="${Math.round(beforeLines * 100)}"`)
  if (afterLines !== undefined && afterLines !== null) spacingParts.push(`w:afterLines="${Math.round(afterLines * 100)}"`)

  const line = spec.spacing?.line ?? spec.lineSpacing
  const lineRule = spec.spacing?.lineRule || (line !== undefined && line !== null && line <= 10 ? 'auto' : null)
  if (line !== undefined && line !== null) {
    const raw = lineRule === 'auto' ? Math.round(line * 240) : pointsToTwips(line)
    spacingParts.push(`w:line="${raw}"`)
    if (lineRule) spacingParts.push(`w:lineRule="${lineRule}"`)
  }
  if (spacingParts.length > 0) parts.push(`<w:spacing ${spacingParts.join(' ')}/>`)

  const indParts = []
  const ind = spec.indents || {}
  if (ind.left !== undefined && ind.left !== null) indParts.push(`w:left="${pointsToTwips(ind.left)}"`)
  if (ind.right !== undefined && ind.right !== null) indParts.push(`w:right="${pointsToTwips(ind.right)}"`)
  if (ind.firstLine !== undefined && ind.firstLine !== null) indParts.push(`w:firstLine="${pointsToTwips(ind.firstLine)}"`)
  if (ind.hanging !== undefined && ind.hanging !== null) indParts.push(`w:hanging="${pointsToTwips(ind.hanging)}"`)
  if (indParts.length > 0) parts.push(`<w:ind ${indParts.join(' ')}/>`)

  const alignment = normalizeAlignment(spec.alignment || spec.align)
  if (alignment) parts.push(`<w:jc w:val="${escapeXml(alignment)}"/>`)

  if (spec.outlineLevel !== undefined && spec.outlineLevel !== null) {
    parts.push(`<w:outlineLvl w:val="${Number(spec.outlineLevel)}"/>`)
  }

  return parts.join('')
}

/**
 * Build one border edge element for a paragraph or table cell.
 * @param {string} edge
 * @param {object|string} spec - `{ style, sizePoints, color, pointSize }` or a style name
 * @returns {string}
 */
export function buildBorderEdge(edge, spec) {
  const detail = typeof spec === 'string' ? { style: spec } : (spec || {})
  const style = detail.style || 'single'
  const sizePoints = detail.sizePoints ?? detail.widthPoints ?? detail.width ?? null
  const eighths = detail.sizeEighthsOfPoint
    ?? (sizePoints === null ? 4 : Math.max(2, Math.round(sizePoints * 8)))
  const color = normalizeColor(detail.color) || 'auto'
  const space = detail.spacePoints === undefined || detail.spacePoints === null
    ? 0
    : pointsToTwips(detail.spacePoints)
  return `<w:${edge} w:val="${escapeXml(style)}" w:sz="${eighths}" w:space="${space}" w:color="${escapeXml(color.replace(/^#/, ''))}"/>`
}

/**
 * Build a `<w:r>` element from text and normalized formatting.
 * @param {string} text
 * @param {object} [spec]
 * @returns {string}
 */
export function buildRunXml(text, spec = {}) {
  const rPr = buildRunProperties(spec)
  const properties = rPr ? `<w:rPr>${rPr}</w:rPr>` : ''
  return `<w:r>${properties}<w:t xml:space="preserve">${escapeXml(text ?? '')}</w:t></w:r>`
}

/**
 * Build a complete `<w:p>` element from text and normalized formatting.
 *
 * `spec.text` (or `text` as a bare string) is the paragraph content; every
 * other key is formatting. `spec.sections`/`spec.sectPr` are handled by the
 * sections module, not here.
 *
 * @param {string|object} spec
 * @param {object} [context]
 * @returns {string}
 */
export function buildParagraphXml(spec, context = {}) {
  const detail = typeof spec === 'string' ? { text: spec } : (spec || {})
  const pPr = buildParagraphProperties(detail, context)
  const properties = pPr ? `<w:pPr>${pPr}</w:pPr>` : ''

  const runs = []
  if (detail.text !== undefined && detail.text !== null && detail.text !== '') {
    runs.push(buildRunXml(detail.text, detail))
  }
  for (const run of detail.runs || []) {
    runs.push(buildRunXml(typeof run === 'string' ? run : run.text, typeof run === 'string' ? {} : run))
  }

  return `<w:p>${properties}${runs.join('')}</w:p>`
}

/**
 * Ensure `word/styles.xml` defines the styles the engine is about to use.
 *
 * The stock R7 template ships only `Normal` and a handful of utility styles:
 * the heading styles are latent, which means a `<w:pStyle w:val="Heading1"/>`
 * has no definition to resolve against. Appending the missing definitions
 * (never replacing an existing one) is what makes an engine-authored document
 * look like a document rather than a wall of body text.
 *
 * @param {string} stylesXml
 * @param {string[]} styleIds
 * @returns {{xml: string, added: string[]}}
 */
export function ensureStyleDefinitions(stylesXml, styleIds = []) {
  if (typeof stylesXml !== 'string' || stylesXml.length === 0) {
    return { xml: stylesXml, added: [] }
  }
  const existing = parseStyleDefinitions(stylesXml)
  // R7's own styles.xml names its default paragraph style `a`, Word names it
  // `Normal`; a new style must inherit from whichever one this package has.
  const normalStyleId = findStyleIdByName(existing, 'Normal') || 'Normal'
  const added = []
  let xml = stylesXml

  for (const styleId of styleIds) {
    if (existing.has(styleId) || added.includes(styleId)) continue
    const definition = styleDefinitionXml(styleId, { normalStyleId })
    if (!definition) continue
    if (xml.includes('</w:styles>')) {
      xml = xml.replace('</w:styles>', `${definition}</w:styles>`)
    } else {
      break
    }
    added.push(styleId)
  }

  return { xml, added }
}

/**
 * Find the style id whose display name matches, case-insensitively.
 * @param {Map<string, object>} styles
 * @param {string} name
 * @returns {string|null}
 */
export function findStyleIdByName(styles, name) {
  if (!styles) return null
  const wanted = String(name).toLowerCase()
  for (const [id, def] of styles) {
    if (String(def.name).toLowerCase() === wanted) return id
  }
  return null
}

/**
 * A self-contained definition for a known built-in style.
 * @param {string} styleId
 * @param {object} [context]
 * @param {string} [context.normalStyleId]
 * @returns {string|null}
 */
export function styleDefinitionXml(styleId, context = {}) {
  const base = context.normalStyleId || 'Normal'
  const heading = /^Heading([1-6])$/.exec(styleId)
  if (heading) {
    const level = Number(heading[1])
    const sizes = [32, 28, 26, 24, 22, 22]
    const before = level === 1 ? 240 : 200
    return '<w:style w:type="paragraph" w:styleId="' + styleId + '">'
      + `<w:name w:val="heading ${level}"/><w:basedOn w:val="${base}"/><w:next w:val="${base}"/>`
      + `<w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/>`
      + `<w:spacing w:before="${before}" w:after="80" w:line="259" w:lineRule="auto"/>`
      + `<w:outlineLvl w:val="${level - 1}"/></w:pPr>`
      + `<w:rPr><w:b/><w:sz w:val="${sizes[level - 1]}"/><w:szCs w:val="${sizes[level - 1]}"/></w:rPr></w:style>`
  }
  if (styleId === 'Hyperlink') {
    return '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>'
      + '<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
  }
  if (styleId === 'ListParagraph') {
    return '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
      + `<w:basedOn w:val="${base}"/><w:uiPriority w:val="34"/><w:qFormat/>`
      + '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
  }
  if (styleId === 'Title') {
    return '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
      + `<w:basedOn w:val="${base}"/><w:next w:val="${base}"/><w:uiPriority w:val="10"/><w:qFormat/>`
      + '<w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr>'
      + '<w:rPr><w:b/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr></w:style>'
  }
  if (styleId === 'Subtitle') {
    return '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/>'
      + `<w:basedOn w:val="${base}"/><w:next w:val="${base}"/><w:uiPriority w:val="11"/><w:qFormat/>`
      + '<w:pPr><w:spacing w:after="160"/><w:jc w:val="center"/></w:pPr>'
      + '<w:rPr><w:i/><w:color w:val="595959"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>'
  }
  return null
}
