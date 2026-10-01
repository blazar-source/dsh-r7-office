/**
 * XLSX style model: `xl/styles.xml`.
 *
 * The whole point of this module is that it never assumes anything about the
 * numbering inside a stylesheet it did not create. A workbook authored by R7,
 * Excel or LibreOffice can have any number of fonts, fills, borders and cell
 * formats, in any order, with or without optional parts. Everything here is
 * therefore parse-then-append:
 *
 *   - existing entries keep their exact indices, because every cell in the
 *     workbook refers to them by number and rebuilding the tables silently
 *     repaints the document;
 *   - new entries are appended and deduplicated against the parsed tables;
 *   - an unknown or unparsable part is preserved verbatim rather than dropped.
 *
 * The public surface is deliberately index-oriented (`ensureFont` returns an
 * index) because that is what a worksheet needs to write into `<c s="…">`.
 */

import { escapeXml, unescapeXml, extractElements, getAttribute } from '../shared/xml.js'

/** Number-format ids below this are reserved for built-in formats. */
export const FIRST_CUSTOM_NUMFMT_ID = 164

/** Built-in number-format ids used by the friendly format names. */
export const BUILTIN_NUMFMT = {
  general: 0,
  integer: 1,          // 0
  decimal: 2,          // 0.00
  thousands: 3,        // #,##0
  thousandsDecimal: 4, // #,##0.00
  percent: 9,          // 0%
  percentDecimal: 10,  // 0.00%
  date: 14,            // locale date
  datetime: 22,        // locale date + time
  time: 21,
  text: 49
}

/** Border line styles accepted by OOXML, thinnest to thickest. */
export const BORDER_STYLES = new Set([
  'none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double',
  'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot',
  'mediumDashDotDot', 'slantDashDot'
])

/** A border edge width in 1/8 pt, used to express a requested thickness. */
export const BORDER_WIDTH_TO_STYLE = new Map([
  [1, 'thin'],
  [2, 'medium'],
  [3, 'thick']
])

const BORDER_ORDER = ['left', 'right', 'top', 'bottom', 'diagonal']

/**
 * Normalize a colour to the ARGB form OOXML expects.
 * @param {string|null|undefined} value - `#RRGGBB`, `RRGGBB`, `AARRGGBB` or `#AARRGGBB`.
 * @returns {string|null}
 */
export function normalizeArgb(value) {
  if (!value) return null
  let hex = String(value).trim().replace(/^#/, '').toUpperCase()
  if (!/^[0-9A-F]{6}$/.test(hex) && !/^[0-9A-F]{8}$/.test(hex)) {
    throw new Error(`Invalid colour: "${value}". Use #RRGGBB or #AARRGGBB.`)
  }
  if (hex.length === 6) hex = `FF${hex}`
  return hex
}

/**
 * Convert a JS Date or ISO 8601 string to an Excel serial day number.
 *
 * Excel's epoch is 1899-12-30 (day 1), which is what makes 1900-02-29 a
 * phantom leap day that every spreadsheet still honours.
 *
 * @param {string|Date} value
 * @returns {number}
 */
export function dateToSerial(value) {
  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) {
    throw new Error(`Invalid date: "${value}". Use an ISO 8601 date or datetime string.`)
  }
  const epoch = Date.UTC(1899, 11, 30)
  const utc = Date.UTC(
    date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(),
    date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds()
  )
  return (utc - epoch) / 86400000
}

/**
 * Inverse of {@link dateToSerial}.
 * @param {number} serial
 * @returns {{date: string, datetime: string}}
 */
export function serialToDate(serial) {
  const epoch = Date.UTC(1899, 11, 30)
  const ms = epoch + Math.round(serial * 86400000)
  const d = new Date(ms)
  const pad = (n) => String(n).padStart(2, '0')
  const date = `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`
  const datetime = `${date}T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`
  return { date, datetime }
}

/**
 * Resolve a friendly number-format request into a format code.
 *
 * @param {object} spec
 * @param {string} [spec.type] - integer | decimal | currency | percent | date | datetime | time | text | custom | general
 * @param {number} [spec.decimals]
 * @param {string} [spec.symbol] - currency symbol
 * @param {string} [spec.code] - explicit format code (`type: 'custom'`)
 * @param {boolean} [spec.thousands]
 * @returns {{code: string|null, builtinId: number|null}}
 */
export function resolveNumberFormat(spec = {}) {
  const type = String(spec.type || '').toLowerCase()
  const decimals = Number.isInteger(spec.decimals) ? spec.decimals : null
  const symbol = spec.symbol === undefined ? '₽' : spec.symbol
  const thousands = spec.thousands === true

  switch (type) {
    case 'general':
      return { code: null, builtinId: 0 }
    case 'integer':
      return { code: thousands ? '#,##0' : '0', builtinId: null }
    case 'decimal': {
      const d = decimals === null ? 2 : decimals
      return { code: `${thousands ? '#,##0' : '0'}.${'0'.repeat(d)}`, builtinId: null }
    }
    case 'currency': {
      const d = decimals === null ? 2 : decimals
      const body = `${thousands === false ? '0' : '#,##0'}.${'0'.repeat(d)}`
      // The symbol is a literal, so it is quoted to keep it out of the format grammar.
      const suffix = symbol ? `"${symbol}"` : ''
      return { code: `${body} ${suffix}`.trim(), builtinId: null }
    }
    case 'percent': {
      const d = decimals === null ? 0 : decimals
      return { code: d === 0 ? '0%' : `0.${'0'.repeat(d)}%`, builtinId: null }
    }
    case 'date':
      return { code: 'DD.MM.YYYY', builtinId: null }
    case 'datetime':
      return { code: 'DD.MM.YYYY HH:MM:SS', builtinId: null }
    case 'time':
      return { code: 'HH:MM:SS', builtinId: null }
    case 'text':
      return { code: '@', builtinId: null }
    case 'custom':
      if (!spec.code) throw new Error('A custom number format requires a "code" value.')
      return { code: String(spec.code), builtinId: null }
    default:
      throw new Error(
        `Unsupported number format type: "${spec.type}". `
        + 'Use integer, decimal, currency, percent, date, datetime, time, text, custom or general.'
      )
  }
}

/** Compare two plain objects by their JSON shape, ignoring key order. */
function sameShape(a, b) {
  const norm = (o) => JSON.stringify(o, Object.keys(o || {}).sort())
  return norm(a) === norm(b)
}

/** Read a `<font>` element into a normalized description. */
function parseFont(xml) {
  if (!xml) return { raw: '' }
  const color = extractElements(xml, 'color')[0]
  const u = extractElements(xml, 'u')[0]
  return {
    bold: /<b\s*\/>|<b\s+val="(1|true)"/.test(xml),
    italic: /<i\s*\/>|<i\s+val="(1|true)"/.test(xml),
    underline: u ? (getAttribute(u.outerXml, 'val') || 'single') : false,
    strike: /<strike\s*\/>/.test(xml),
    size: Number(getAttribute(extractElements(xml, 'sz')[0]?.outerXml || '', 'val')) || null,
    name: getAttribute(extractElements(xml, 'name')[0]?.outerXml || '', 'val') || null,
    family: getAttribute(extractElements(xml, 'family')[0]?.outerXml || '', 'val') || null,
    scheme: getAttribute(extractElements(xml, 'scheme')[0]?.outerXml || '', 'val') || null,
    color: color
      ? {
        rgb: getAttribute(color.outerXml, 'rgb'),
        theme: getAttribute(color.outerXml, 'theme'),
        indexed: getAttribute(color.outerXml, 'indexed'),
        tint: getAttribute(color.outerXml, 'tint')
      }
      : null,
    raw: xml
  }
}

/** Read a `<fill>` element into a normalized description. */
function parseFill(xml) {
  if (!xml) return { patternType: 'none', raw: '' }
  const patternType = getAttribute(xml, 'patternType') || 'none'
  const fg = extractElements(xml, 'fgColor')[0]
  const bg = extractElements(xml, 'bgColor')[0]
  return {
    patternType,
    fgColor: fg ? {
      rgb: getAttribute(fg.outerXml, 'rgb'),
      theme: getAttribute(fg.outerXml, 'theme'),
      indexed: getAttribute(fg.outerXml, 'indexed')
    } : null,
    bgColor: bg ? {
      rgb: getAttribute(bg.outerXml, 'rgb'),
      theme: getAttribute(bg.outerXml, 'theme'),
      indexed: getAttribute(bg.outerXml, 'indexed')
    } : null,
    raw: xml
  }
}

/** Read a border edge element. */
function parseBorderEdge(element) {
  if (!element) return null
  const style = getAttribute(element.outerXml, 'style')
  if (!style || style === 'none') return null
  const color = extractElements(element.outerXml, 'color')[0]
  return {
    style,
    color: color
      ? {
        rgb: getAttribute(color.outerXml, 'rgb'),
        theme: getAttribute(color.outerXml, 'theme'),
        indexed: getAttribute(color.outerXml, 'indexed')
      }
      : null
  }
}

/** Read a `<border>` element into a normalized description. */
function parseBorder(xml) {
  if (!xml) return { raw: '' }
  const out = { raw: xml }
  for (const edge of BORDER_ORDER) {
    const element = extractElements(xml, edge)[0]
    out[edge] = edge === 'diagonal' ? undefined : parseBorderEdge(element)
  }
  out.diagonalUp = /diagonalUp="1"/.test(xml)
  out.diagonalDown = /diagonalDown="1"/.test(xml)
  return out
}

/** Read an `<xf>` element into a normalized description. */
function parseXf(xml) {
  if (!xml) return null
  const alignment = extractElements(xml, 'alignment')[0]
  const protection = extractElements(xml, 'protection')[0]
  return {
    numFmtId: Number(getAttribute(xml, 'numFmtId') || 0),
    fontId: Number(getAttribute(xml, 'fontId') || 0),
    fillId: Number(getAttribute(xml, 'fillId') || 0),
    borderId: Number(getAttribute(xml, 'borderId') || 0),
    xfId: getAttribute(xml, 'xfId'),
    alignment: alignment ? {
      horizontal: getAttribute(alignment.outerXml, 'horizontal'),
      vertical: getAttribute(alignment.outerXml, 'vertical'),
      wrapText: ['1', 'true'].includes(getAttribute(alignment.outerXml, 'wrapText')),
      textRotation: getAttribute(alignment.outerXml, 'textRotation'),
      indent: getAttribute(alignment.outerXml, 'indent'),
      shrinkToFit: ['1', 'true'].includes(getAttribute(alignment.outerXml, 'shrinkToFit'))
    } : null,
    protection: protection ? protection.outerXml : null,
    raw: xml
  }
}

/**
 * A parsed `xl/styles.xml` that can be extended without disturbing what is
 * already there.
 */
export class Stylesheet {
  constructor(xml = null) {
    this.prefix = ''
    this.suffix = ''
    this.numFmts = []      // { id, code, raw }
    this.fonts = []
    this.fills = []
    this.borders = []
    this.cellStyleXfs = []
    this.cellXfs = []
    this.preserved = []    // parts we do not model, kept verbatim
    /** Original container opening tags, so extra attributes survive. */
    this.containerTags = {}
    /** Original document, returned verbatim while nothing has changed. */
    this.original = xml
    this.dirty = false

    if (xml) this.parse(xml)
    else this._seedDefaults()
  }

  /** Remember the opening tag of a container, minus its count attribute. */
  _rememberContainer(xml, tag) {
    const el = extractElements(xml, tag)[0]
    if (!el) return
    const open = el.outerXml.match(/^<[^>]*>/)
    if (!open) return
    // Keep everything up to (but not including) the closing angle bracket, so
    // _containerOpen can append attributes before it.
    this.containerTags[tag] = open[0]
      .replace(/>$/, '')
      .replace(/\s*count="[^"]*"/, '')
      .replace(/\s*uniqueCount="[^"]*"/, '')
  }

  /** Emit a container opening tag carrying the current element count. */
  _containerOpen(tag, count) {
    const base = this.containerTags[tag] || `<${tag}`
    return `${base} count="${count}">`
  }

  /** Minimal but valid stylesheet, matching what R7 ships. */
  _seedDefaults() {
    this.prefix = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    this.suffix = '</styleSheet>'
    this.fonts.push(parseFont('<font><sz val="11"/><color theme="1"/><name val="Liberation Sans"/><family val="2"/><scheme val="minor"/></font>'))
    this.fills.push(parseFill('<fill><patternFill patternType="none"/></fill>'))
    this.fills.push(parseFill('<fill><patternFill patternType="gray125"/></fill>'))
    this.borders.push(parseBorder('<border><left/><right/><top/><bottom/><diagonal/></border>'))
    this.cellStyleXfs.push(parseXf('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'))
    this.cellXfs.push(parseXf('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'))
  }

  /**
   * Parse an existing stylesheet, keeping every part we do not model.
   * @param {string} xml
   */
  parse(xml) {
    const sheetTag = xml.match(/<styleSheet[^>]*>/)
    if (!sheetTag) throw new Error('Invalid XLSX: xl/styles.xml has no <styleSheet> root')

    this.prefix = xml.slice(0, sheetTag.index + sheetTag[0].length)
    this.suffix = '</styleSheet>'

    const numFmtsXml = extractElements(xml, 'numFmts')[0]
    if (numFmtsXml) {
      for (const el of extractElements(numFmtsXml.outerXml, 'numFmt')) {
        this.numFmts.push({
          id: Number(getAttribute(el.outerXml, 'numFmtId')),
          code: unescapeXml(getAttribute(el.outerXml, 'formatCode') || ''),
          raw: el.outerXml
        })
      }
    }

    const fontsXml = extractElements(xml, 'fonts')[0]
    if (fontsXml) {
      const container = fontsXml.outerXml.replace(/^<fonts[^>]*>/, '').replace(/<\/fonts>$/, '')
      for (const el of extractElements(container, 'font')) this.fonts.push(parseFont(el.outerXml))
    }

    const fillsXml = extractElements(xml, 'fills')[0]
    if (fillsXml) {
      const container = fillsXml.outerXml.replace(/^<fills[^>]*>/, '').replace(/<\/fills>$/, '')
      for (const el of extractElements(container, 'fill')) this.fills.push(parseFill(el.outerXml))
    }

    const bordersXml = extractElements(xml, 'borders')[0]
    if (bordersXml) {
      const container = bordersXml.outerXml.replace(/^<borders[^>]*>/, '').replace(/<\/borders>$/, '')
      for (const el of extractElements(container, 'border')) this.borders.push(parseBorder(el.outerXml))
    }

    const styleXfs = extractElements(xml, 'cellStyleXfs')[0]
    if (styleXfs) {
      for (const el of extractElements(styleXfs.outerXml, 'xf')) this.cellStyleXfs.push(parseXf(el.outerXml))
    }

    const cellXfs = extractElements(xml, 'cellXfs')[0]
    if (cellXfs) {
      for (const el of extractElements(cellXfs.outerXml, 'xf')) this.cellXfs.push(parseXf(el.outerXml))
    }

    // Anything else (cellStyles, dxfs, tableStyles, colors, extLst) is carried
    // through untouched so a round trip cannot silently drop it.
    for (const tag of ['cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']) {
      const el = extractElements(xml, tag)[0]
      if (el) this.preserved.push(el.outerXml)
    }

    for (const tag of ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs']) {
      this._rememberContainer(xml, tag)
    }
  }

  /** Find or append a font, returning its index. */
  ensureFont(spec) {
    if (!spec || Object.keys(spec).length === 0) return 0
    const existing = this.fonts.findIndex((f) => this._fontMatches(f, spec))
    if (existing !== -1) return existing
    this.dirty = true
    this.fonts.push(this._buildFont(spec))
    return this.fonts.length - 1
  }

  /**
   * A font spec suitable for writing, as opposed to {@link describeFont} which
   * is for reporting. Theme and indexed colours survive here, because a font
   * that inherits its colour from the theme must not be flattened to a literal.
   */
  _fontSpecForWrite(index) {
    const font = this.fonts[index]
    if (!font) return {}
    const spec = {}
    if (font.bold) spec.bold = true
    if (font.italic) spec.italic = true
    if (font.underline) spec.underline = font.underline
    if (font.strike) spec.strike = true
    if (font.size !== null) spec.size = font.size
    if (font.name !== null) spec.family = font.name
    if (font.color) spec.color = { ...font.color }
    return spec
  }

  /** Compare a colour specification, which may be a string or a descriptor. */
  _colorMatches(font, spec) {
    const want = spec.color
    const have = font.color || null
    if (want === undefined) return true
    if (want === null) return !have
    if (typeof want === 'string') return (have?.rgb || null) === normalizeArgb(want)
    if (want.rgb !== undefined && want.rgb !== null) return (have?.rgb || null) === normalizeArgb(want.rgb)
    if (want.theme !== undefined && want.theme !== null) return String(have?.theme ?? '') === String(want.theme)
    if (want.indexed !== undefined && want.indexed !== null) return String(have?.indexed ?? '') === String(want.indexed)
    return true
  }

  _fontMatches(font, spec) {
    if (spec.bold !== undefined && font.bold !== Boolean(spec.bold)) return false
    if (spec.italic !== undefined && font.italic !== Boolean(spec.italic)) return false
    if (spec.underline !== undefined) {
      const wanted = spec.underline === true ? 'single' : (spec.underline || false)
      if ((font.underline || false) !== wanted) return false
    }
    if (spec.strike !== undefined && font.strike !== Boolean(spec.strike)) return false
    if (spec.size !== undefined && Number(font.size) !== Number(spec.size)) return false
    if (spec.family !== undefined && font.name !== spec.family) return false
    return this._colorMatches(font, spec)
  }

  /** Render a colour descriptor as an OOXML `<color …/>` element. */
  _colorElement(color) {
    if (!color) return ''
    if (typeof color === 'string') return `<color rgb="${normalizeArgb(color)}"/>`
    if (color.rgb) return `<color rgb="${normalizeArgb(color.rgb)}"/>`
    if (color.theme !== undefined && color.theme !== null) {
      const tint = color.tint !== undefined && color.tint !== null ? ` tint="${color.tint}"` : ''
      return `<color theme="${color.theme}"${tint}/>`
    }
    if (color.indexed !== undefined && color.indexed !== null) {
      return `<color indexed="${color.indexed}"/>`
    }
    return ''
  }

  _buildFont(spec) {
    const parts = []
    if (spec.bold) parts.push('<b/>')
    if (spec.italic) parts.push('<i/>')
    if (spec.underline) {
      parts.push(spec.underline === true || spec.underline === 'single'
        ? '<u/>'
        : `<u val="${escapeXml(String(spec.underline))}"/>`)
    }
    if (spec.strike) parts.push('<strike/>')
    if (spec.size !== undefined && spec.size !== null) parts.push(`<sz val="${Number(spec.size)}"/>`)
    parts.push(this._colorElement(spec.color))
    if (spec.family !== undefined && spec.family !== null) {
      parts.push(`<name val="${escapeXml(String(spec.family))}"/>`)
      parts.push('<family val="2"/>')
    }
    const xml = `<font>${parts.join('')}</font>`
    return parseFont(xml)
  }

  /** Find or append a solid fill, returning its index. */
  ensureFill(spec) {
    if (!spec || spec.color === undefined || spec.color === null) return 0
    const wanted = {
      patternType: spec.patternType || 'solid',
      fgColor: { rgb: normalizeArgb(spec.color), theme: null, indexed: null },
      bgColor: { rgb: null, theme: null, indexed: '64' }
    }
    const existing = this.fills.findIndex((f) => sameShape(
      { p: f.patternType, fg: f.fgColor?.rgb || null, bg: f.bgColor?.indexed || f.bgColor?.rgb || null },
      { p: wanted.patternType, fg: wanted.fgColor.rgb, bg: wanted.bgColor.indexed }
    ))
    if (existing !== -1) return existing

    const xml = `<fill><patternFill patternType="${wanted.patternType}">`
      + `<fgColor rgb="${wanted.fgColor.rgb}"/>`
      + `<bgColor indexed="64"/>`
      + '</patternFill></fill>'
    this.dirty = true
    this.fills.push(parseFill(xml))
    return this.fills.length - 1
  }

  /** Find or append a border, returning its index. */
  ensureBorder(spec) {
    if (!spec || Object.keys(spec).length === 0) return 0
    const edges = {}
    for (const edge of ['left', 'right', 'top', 'bottom']) {
      const value = spec[edge]
      if (!value) continue
      const style = typeof value === 'string' ? value : (value.style || 'thin')
      if (!BORDER_STYLES.has(style) || style === 'none') continue
      edges[edge] = {
        style,
        color: normalizeArgb(typeof value === 'object' ? (value.color || '#000000') : (spec.color || '#000000'))
      }
    }
    if (Object.keys(edges).length === 0) return 0

    const existing = this.borders.findIndex((b) => sameShape(
      Object.fromEntries(['left', 'right', 'top', 'bottom'].map((e) => [
        e, b[e] ? { style: b[e].style, rgb: b[e].color?.rgb || null } : null
      ])),
      Object.fromEntries(['left', 'right', 'top', 'bottom'].map((e) => [
        e, edges[e] ? { style: edges[e].style, rgb: edges[e].color } : null
      ]))
    ))
    if (existing !== -1) return existing

    const parts = ['<border>']
    for (const edge of ['left', 'right', 'top', 'bottom']) {
      parts.push(edges[edge]
        ? `<${edge} style="${edges[edge].style}"><color rgb="${edges[edge].color}"/></${edge}>`
        : `<${edge}/>`)
    }
    parts.push('<diagonal/></border>')
    this.dirty = true
    this.borders.push(parseBorder(parts.join('')))
    return this.borders.length - 1
  }

  /**
   * Find or append a number format, returning its id.
   * @param {object} spec - resolved format request.
   * @returns {number}
   */
  ensureNumFmt(spec) {
    const { code, builtinId } = resolveNumberFormat(spec)
    if (builtinId !== null && code === null) return builtinId
    if (code === null) return 0

    // Reuse an existing declaration with the same code, built-in or custom.
    const existingBuiltin = Object.entries(BUILTIN_NUMFMT).find(([, id]) => id !== undefined
      && this.numFmts.some((n) => n.id === id && n.code === code))
    if (existingBuiltin) return existingBuiltin[1]

    const declared = this.numFmts.find((n) => n.code === code)
    if (declared) return declared.id

    const used = this.numFmts.map((n) => n.id).filter((n) => n >= FIRST_CUSTOM_NUMFMT_ID)
    const id = used.length > 0 ? Math.max(...used, FIRST_CUSTOM_NUMFMT_ID - 1) + 1 : FIRST_CUSTOM_NUMFMT_ID
    this.dirty = true
    this.numFmts.push({ id, code, raw: '' })
    return id
  }

  /**
   * Find or append a cell format (`<xf>` in `cellXfs`), returning its index.
   *
   * When `base` is given, the new format inherits every property of that cell
   * format and only the requested properties are overridden — which is what
   * makes "make this cell bold" keep its borders, fill and number format.
   *
   * @param {object} spec
   * @param {number} [spec.base] - existing cellXfs index to start from
   * @param {object} [spec.font] - font spec (merged with the base font)
   * @param {object} [spec.fill]
   * @param {object} [spec.border]
   * @param {object} [spec.numberFormat]
   * @param {object} [spec.alignment]
   * @returns {number}
   */
  ensureCellXf(spec = {}) {
    const baseIndex = Number.isInteger(spec.base) ? spec.base : 0
    const base = this.cellXfs[baseIndex] || this.cellXfs[0] || parseXf('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')

    let fontId = base.fontId
    if (spec.font) {
      const merged = { ...this._fontSpecForWrite(base.fontId), ...spec.font }
      fontId = this.ensureFont(merged)
    }

    let fillId = base.fillId
    if (spec.fill) {
      fillId = spec.fill.color === null || spec.fill.color === undefined
        ? base.fillId
        : this.ensureFill(spec.fill)
    }

    let borderId = base.borderId
    if (spec.border) {
      const mergedEdges = {}
      for (const edge of ['left', 'right', 'top', 'bottom']) {
        if (spec.border[edge] !== undefined) mergedEdges[edge] = spec.border[edge]
        else if (base[edge]) mergedEdges[edge] = base[edge]
      }
      // A border spec that only clears edges must still reference a border entry.
      if (Object.keys(mergedEdges).length > 0) {
        borderId = this.ensureBorder({ ...spec.border, ...mergedEdges })
      } else {
        const current = this.borders[base.borderId]
        if (current) {
          const hasEdge = ['left', 'right', 'top', 'bottom'].some((e) => current[e])
          borderId = hasEdge ? this.ensureBorder({ ...this._borderSpec(current), ...spec.border }) : base.borderId
        }
      }
    }

    let numFmtId = base.numFmtId
    if (spec.numberFormat) numFmtId = this.ensureNumFmt(spec.numberFormat)

    let alignment = base.alignment
    if (spec.alignment) {
      alignment = { ...(base.alignment || {}), ...spec.alignment }
      for (const [k, v] of Object.entries(alignment)) {
        if (v === undefined || v === null) delete alignment[k]
      }
    }

    const candidate = { numFmtId, fontId, fillId, borderId, xfId: base.xfId ?? '0', alignment }

    const existing = this.cellXfs.findIndex((x) => sameShape(
      {
        n: x.numFmtId, f: x.fontId, fi: x.fillId, b: x.borderId,
        a: x.alignment || null
      },
      {
        n: candidate.numFmtId, f: candidate.fontId, fi: candidate.fillId, b: candidate.borderId,
        a: candidate.alignment || null
      }
    ))
    if (existing !== -1) return existing

    const parts = [
      `<xf numFmtId="${candidate.numFmtId}"`,
      ` fontId="${candidate.fontId}"`,
      ` fillId="${candidate.fillId}"`,
      ` borderId="${candidate.borderId}"`,
      ` xfId="${candidate.xfId}"`
    ]
    if (candidate.numFmtId !== base.numFmtId || spec.numberFormat) parts.push(' applyNumberFormat="1"')
    if (fontId !== base.fontId || spec.font) parts.push(' applyFont="1"')
    if (fillId !== base.fillId || spec.fill) parts.push(' applyFill="1"')
    if (borderId !== base.borderId || spec.border) parts.push(' applyBorder="1"')

    let inner = ''
    if (alignment && Object.keys(alignment).length > 0) {
      parts.push(' applyAlignment="1"')
      const attrs = []
      if (alignment.horizontal) attrs.push(`horizontal="${escapeXml(alignment.horizontal)}"`)
      if (alignment.vertical) attrs.push(`vertical="${escapeXml(alignment.vertical)}"`)
      if (alignment.wrapText) attrs.push('wrapText="1"')
      if (alignment.textRotation !== undefined) attrs.push(`textRotation="${alignment.textRotation}"`)
      if (alignment.indent !== undefined) attrs.push(`indent="${alignment.indent}"`)
      if (alignment.shrinkToFit) attrs.push('shrinkToFit="1"')
      inner = `<alignment ${attrs.join(' ')}/>`
    }

    const xml = inner
      ? `${parts.join('')}>${inner}</xf>`
      : `${parts.join('')}/>`
    this.dirty = true
    this.cellXfs.push(parseXf(xml))
    return this.cellXfs.length - 1
  }

  _borderSpec(border) {
    const out = {}
    for (const edge of ['left', 'right', 'top', 'bottom']) {
      if (border[edge]) out[edge] = { style: border[edge].style, color: border[edge].color?.rgb || null }
    }
    return out
  }

  /** Normalized font description for a font index. */
  describeFont(index) {
    const font = this.fonts[index]
    if (!font) return { bold: false, italic: false, underline: false, strike: false, size: null, family: null, color: null }
    return {
      family: font.name,
      size: font.size,
      bold: Boolean(font.bold),
      italic: Boolean(font.italic),
      underline: font.underline || false,
      strike: Boolean(font.strike),
      color: font.color?.rgb || (font.color?.theme !== undefined && font.color?.theme !== null ? `theme:${font.color.theme}` : null)
    }
  }

  /** Normalized fill description for a fill index. */
  describeFill(index) {
    const fill = this.fills[index]
    if (!fill || fill.patternType === 'none') return null
    return {
      patternType: fill.patternType,
      color: fill.fgColor?.rgb || (fill.fgColor?.theme !== undefined && fill.fgColor?.theme !== null
        ? `theme:${fill.fgColor.theme}`
        : null)
    }
  }

  /** Normalized border description for a border index. */
  describeBorder(index) {
    const border = this.borders[index]
    if (!border) return null
    const out = {}
    for (const edge of ['left', 'right', 'top', 'bottom']) {
      if (border[edge]) {
        out[edge] = { style: border[edge].style, color: border[edge].color?.rgb || null }
      }
    }
    return Object.keys(out).length > 0 ? out : null
  }

  /** The format code for a number-format id, or a named built-in. */
  describeNumberFormat(id) {
    const declared = this.numFmts.find((n) => n.id === id)
    if (declared) return declared.code
    const known = Object.entries(BUILTIN_NUMFMT).find(([, v]) => v === id)
    return known ? known[0] : `builtin:${id}`
  }

  /** Fully normalized description of a cell format. */
  describeCellXf(index) {
    const xf = this.cellXfs[index]
    if (!xf) return null
    return {
      index,
      font: this.describeFont(xf.fontId),
      fill: this.describeFill(xf.fillId),
      border: this.describeBorder(xf.borderId),
      alignment: xf.alignment && Object.keys(xf.alignment).length > 0 ? xf.alignment : null,
      numberFormat: this.describeNumberFormat(xf.numFmtId),
      numberFormatId: xf.numFmtId
    }
  }

  /** Serialize back to `xl/styles.xml`. */
  toXml() {
    // Nothing was added: return the document exactly as it arrived. A workbook
    // whose styles this tool never touched must not be rewritten at all.
    if (!this.dirty && this.original) return this.original

    const parts = [this.prefix]

    const customNumFmts = this.numFmts.filter((n) => n.id >= FIRST_CUSTOM_NUMFMT_ID)
    if (customNumFmts.length > 0) {
      parts.push(this._containerOpen('numFmts', customNumFmts.length))
      for (const n of customNumFmts) {
        parts.push(`<numFmt numFmtId="${n.id}" formatCode="${escapeXml(n.code)}"/>`)
      }
      parts.push('</numFmts>')
    }

    parts.push(this._containerOpen('fonts', this.fonts.length))
    for (const f of this.fonts) parts.push(f.raw)
    parts.push('</fonts>')

    parts.push(this._containerOpen('fills', this.fills.length))
    for (const f of this.fills) parts.push(f.raw)
    parts.push('</fills>')

    parts.push(this._containerOpen('borders', this.borders.length))
    for (const b of this.borders) parts.push(b.raw)
    parts.push('</borders>')

    if (this.cellStyleXfs.length > 0) {
      parts.push(this._containerOpen('cellStyleXfs', this.cellStyleXfs.length))
      for (const x of this.cellStyleXfs) parts.push(x.raw)
      parts.push('</cellStyleXfs>')
    }

    parts.push(this._containerOpen('cellXfs', this.cellXfs.length))
    for (const x of this.cellXfs) parts.push(x.raw)
    parts.push('</cellXfs>')

    for (const p of this.preserved) parts.push(p)
    parts.push(this.suffix)
    return parts.join('')
  }

  /** Number of cell formats, i.e. the valid range of `<c s="…">`. */
  get cellFormatCount() {
    return this.cellXfs.length
  }
}
