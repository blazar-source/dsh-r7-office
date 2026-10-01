import { escapeXml, extractElements } from '../shared/xml.js'
import {
  toEmu,
  toArgb,
  colorWithTransparency,
  srgbClrElement,
  xmlAttr,
  withAttribute,
  firstElement,
  findElement,
  removeAll,
  setChild
} from './pptx-util.js'

/**
 * OOXML builders and patchers for the PPTX engine.
 *
 * Two families live here:
 *
 *  - `build*` emits a fresh fragment (a `<p:sp>`, a `<p:txBody>`, a `<p:pic>`),
 *    used when the engine creates something.
 *  - `patch*` edits markup that is already on disk, used when the engine
 *    changes something that exists. Patching the author's own bytes is what
 *    keeps gradient fills, effects, hyperlinks and unknown extensions alive
 *    through a restyle.
 */

// ------------------------------------------------------------- schema orders

/** Child order of `<a:rPr>` as the DrawingML schema requires it. */
export const RPR_ORDER = [
  'a:ln', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill',
  'a:effectLst', 'a:effectDag', 'a:highlight', 'a:uLnTx', 'a:uLn', 'a:uFillTx', 'a:uFill',
  'a:latin', 'a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst'
]

/** Child order of `<a:pPr>`. */
export const PPR_ORDER = [
  'a:lnSpc', 'a:spcBef', 'a:spcAft', 'a:buClrTx', 'a:buClr', 'a:buSzTx', 'a:buSzPct',
  'a:buSzPts', 'a:buFontTx', 'a:buFont', 'a:buNone', 'a:buAutoNum', 'a:buChar',
  'a:tabLst', 'a:defRPr', 'a:extLst'
]

/** Child order of `<p:spPr>`. */
export const SPPR_ORDER = [
  'a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill',
  'a:pattFill', 'a:grpFill', 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'
]

/** Child order of `<p:txBody>`. */
export const TXBODY_ORDER = ['a:bodyPr', 'a:lstStyle', 'a:p']

/** Child order of `<a:bodyPr>`. */
export const BODYPR_ORDER = [
  'a:prstTxWarp', 'a:noAutofit', 'a:normAutofit', 'a:spAutoFit',
  'a:scene3d', 'a:sp3d', 'a:flatTx', 'a:extLst'
]

/** Child order of `<p:sp>`. */
const SP_ORDER = ['p:nvSpPr', 'p:spPr', 'p:txBody', 'p:style', 'p:extLst']

/** Child order of `<p:pic>`. */
const PIC_ORDER = ['p:nvPicPr', 'p:blipFill', 'p:spPr', 'p:style', 'p:extLst']

// ------------------------------------------------------------------ vocabulary

const UNDERLINE_VALUES = new Set([
  'none', 'sng', 'dbl', 'heavy', 'dotted', 'dottedHeavy', 'dash',
  'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy',
  'dotDotDash', 'dotDotDashHeavy', 'wavy', 'wavyHeavy', 'wavyDbl'
])

const UNDERLINE_ALIASES = {
  single: 'sng',
  double: 'dbl',
  heavy: 'heavy',
  dotted: 'dotted',
  'dotted-heavy': 'dottedHeavy',
  dash: 'dash',
  'dash-heavy': 'dashHeavy',
  'dash-long': 'dashLong',
  'dash-long-heavy': 'dashLongHeavy',
  dotdash: 'dotDash',
  'dotdash-heavy': 'dotDashHeavy',
  dotdotdash: 'dotDotDash',
  'dotdotdash-heavy': 'dotDotDashHeavy',
  wavy: 'wavy',
  'wavy-heavy': 'wavyHeavy',
  'wavy-double': 'wavyDbl'
}

const ALIGN_VALUES = new Set(['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist'])
const ALIGN_ALIASES = {
  left: 'l', start: 'l', center: 'ctr', centre: 'ctr', middle: 'ctr',
  right: 'r', end: 'r', justify: 'just', justified: 'just', distributed: 'dist'
}

const ANCHOR_VALUES = new Set(['t', 'ctr', 'b', 'just', 'dist'])
const ANCHOR_ALIASES = {
  top: 't', center: 'ctr', centre: 'ctr', middle: 'ctr',
  bottom: 'b', justify: 'just', justified: 'just', distributed: 'dist'
}

const ARROW_VALUES = new Set(['none', 'triangle', 'stealth', 'arrow', 'diamond', 'oval'])

const PRESET_TYPES = new Set([
  'rect', 'roundRect', 'ellipse', 'line', 'straightConnector1', 'bentConnector2',
  'bentConnector3', 'curvedConnector2', 'curvedConnector3', 'triangle', 'rtTriangle',
  'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'heptagon',
  'octagon', 'decagon', 'dodecagon', 'star4', 'star5', 'star6', 'star7', 'star8',
  'star10', 'star12', 'star16', 'star24', 'star32', 'rightArrow', 'leftArrow',
  'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'bentArrow',
  'uturnArrow', 'chevron', 'homePlate', 'plus', 'mathPlus', 'mathMinus',
  'mathMultiply', 'mathDivide', 'mathEqual', 'flowChartProcess',
  'flowChartDecision', 'flowChartData', 'flowChartPredefinedProcess',
  'flowChartDocument', 'flowChartTerminator', 'cube', 'can', 'cloud',
  'heart', 'sun', 'moon', 'smileyFace', 'arc', 'blockArc', 'pie', 'chord',
  'teardrop', 'frame', 'halfFrame', 'corner', 'diagStripe', 'donut', 'noSmoking',
  'leftBrace', 'rightBrace', 'leftBracket', 'rightBracket', 'round2SameRect',
  'round2DiagRect', 'snip1Rect', 'snip2SameRect', 'snip2DiagRect', 'snipRoundRect'
])

/** The user-facing shape names the engine accepts, mapped to preset geometry. */
export const SHAPE_PRESETS = Object.freeze({
  rectangle: 'rect',
  rect: 'rect',
  square: 'rect',
  'rounded-rectangle': 'roundRect',
  roundedRectangle: 'roundRect',
  roundRect: 'roundRect',
  'rounded rectangle': 'roundRect',
  ellipse: 'ellipse',
  oval: 'ellipse',
  circle: 'ellipse',
  line: 'line',
  'straight-line': 'line',
  arrow: 'rightArrow',
  'arrow-right': 'rightArrow',
  rightArrow: 'rightArrow',
  'arrow-left': 'leftArrow',
  leftArrow: 'leftArrow',
  'arrow-up': 'upArrow',
  upArrow: 'upArrow',
  'arrow-down': 'downArrow',
  downArrow: 'downArrow',
  'arrow-left-right': 'leftRightArrow',
  triangle: 'triangle',
  diamond: 'diamond',
  pentagon: 'pentagon',
  hexagon: 'hexagon',
  octagon: 'octagon',
  star: 'star5',
  'star-4': 'star4',
  'star-5': 'star5',
  'star-6': 'star6',
  chevron: 'chevron',
  plus: 'plus',
  cloud: 'cloud',
  heart: 'heart',
  cylinder: 'can',
  cube: 'cube',
  donut: 'donut',
  pie: 'pie',
  smiley: 'smileyFace',
  parallelogram: 'parallelogram',
  trapezoid: 'trapezoid'
})

/**
 * Resolve a shape type name to a DrawingML preset geometry.
 * @param {string} value
 * @returns {string}
 */
export function resolvePreset(value) {
  if (!value) throw new Error('A shape type is required')
  const raw = String(value).trim()
  if (PRESET_TYPES.has(raw)) return raw
  if (SHAPE_PRESETS[raw]) return SHAPE_PRESETS[raw]
  const lower = raw.toLowerCase()
  if (SHAPE_PRESETS[lower]) return SHAPE_PRESETS[lower]
  if (PRESET_TYPES.has(lower)) return lower
  throw new Error(
    `Unsupported shape type "${value}". Known names: ${Object.keys(SHAPE_PRESETS).join(', ')}.`
  )
}

/** Resolve a paragraph alignment to its DrawingML token. */
export function normalizeAlignment(value) {
  const raw = String(value).trim()
  if (ALIGN_VALUES.has(raw)) return raw
  const mapped = ALIGN_ALIASES[raw.toLowerCase()]
  if (mapped) return mapped
  throw new Error(`Unsupported alignment "${value}". Use left, center, right or justify.`)
}

/** Resolve a vertical anchor to its DrawingML token. */
export function normalizeAnchor(value) {
  const raw = String(value).trim()
  if (ANCHOR_VALUES.has(raw)) return raw
  const mapped = ANCHOR_ALIASES[raw.toLowerCase()]
  if (mapped) return mapped
  throw new Error(`Unsupported vertical alignment "${value}". Use top, center or bottom.`)
}

/** Resolve an underline request to its DrawingML token. */
export function normalizeUnderline(value) {
  if (value === false || value === null) return 'none'
  if (value === true) return 'sng'
  const raw = String(value).trim()
  if (UNDERLINE_VALUES.has(raw)) return raw
  const mapped = UNDERLINE_ALIASES[raw.toLowerCase()]
  if (mapped) return mapped
  throw new Error(
    `Unsupported underline style "${value}". Use true, false, or one of: `
    + `${Object.keys(UNDERLINE_ALIASES).join(', ')}.`
  )
}

/** Resolve a line-end request to its DrawingML token. */
function normalizeArrow(value) {
  if (value === true) return 'triangle'
  const raw = String(value).trim()
  if (ARROW_VALUES.has(raw)) return raw
  throw new Error(
    `Unsupported arrow head "${value}". Use triangle, stealth, arrow, diamond, oval or none.`
  )
}

/**
 * Express the caller's arrow request as a head/tail pair.
 *
 * A straight line points from head to tail, so "an arrow" — the shape a user
 * draws and expects — is a `tailEnd`, which is where the point lands.
 *
 * @param {string|boolean|object} request
 * @returns {{head: string|null, tail: string|null}}
 */
function resolveArrows(request) {
  if (request === undefined || request === null || request === false) return { head: null, tail: null }
  if (request === true) return { head: null, tail: 'triangle' }
  if (typeof request === 'string') {
    const key = request.toLowerCase()
    if (key === 'both' || key === 'double') return { head: 'triangle', tail: 'triangle' }
    if (key === 'end' || key === 'start') return { head: null, tail: 'triangle' }
    return { head: null, tail: normalizeArrow(request) }
  }
  return {
    head: request.head ? normalizeArrow(request.head) : null,
    tail: request.tail ? normalizeArrow(request.tail) : null
  }
}

// --------------------------------------------------------------- run / text

/**
 * Build `<a:rPr>` for one run.
 *
 * @param {object} spec - family, size, bold, italic, underline, color,
 *   transparency, strike, caps, spacing, baseline, highlight, lang, dirty.
 * @returns {string}
 */
export function buildRunProperties(spec = {}) {
  const attrs = []
  const children = []

  attrs.push(['lang', spec.lang || 'ru-RU'])
  if (spec.dirty !== undefined) attrs.push(['dirty', spec.dirty ? '1' : '0'])
  if (spec.size !== undefined && spec.size !== null) {
    const points = Number(spec.size)
    if (!Number.isFinite(points) || points <= 0) {
      throw new Error(`font.size must be a positive point size, got "${spec.size}"`)
    }
    attrs.push(['sz', String(Math.round(points * 100))])
  }
  if (spec.bold !== undefined) attrs.push(['b', spec.bold ? '1' : '0'])
  if (spec.italic !== undefined) attrs.push(['i', spec.italic ? '1' : '0'])
  if (spec.underline !== undefined) attrs.push(['u', normalizeUnderline(spec.underline)])
  if (spec.strike !== undefined) attrs.push(['strike', spec.strike ? 'sngStrike' : 'noStrike'])
  if (spec.caps !== undefined) {
    const caps = String(spec.caps).toLowerCase()
    if (caps === 'small') attrs.push(['cap', 'small'])
    else if (caps === 'all' || caps === 'uppercase') attrs.push(['cap', 'all'])
    else attrs.push(['cap', 'none'])
  }
  if (spec.spacing !== undefined && spec.spacing !== null) {
    // Character spacing is expressed in 1/100 pt.
    const spc = Number(spec.spacing)
    if (!Number.isFinite(spc)) throw new Error(`font.spacing must be a number of points, got "${spec.spacing}"`)
    attrs.push(['spc', String(Math.round(spc * 100))])
  }
  if (spec.baseline !== undefined && spec.baseline !== null) {
    attrs.push(['baseline', String(Math.round(Number(spec.baseline) * 1000))])
  }

  if (spec.color !== undefined && spec.color !== null) {
    children.push(`<a:solidFill>${srgbClrElement(spec.color, spec.transparency)}</a:solidFill>`)
  }
  if (spec.highlight) {
    children.push(`<a:highlight>${srgbClrElement(spec.highlight)}</a:highlight>`)
  }
  if (spec.family) children.push(`<a:latin typeface="${xmlAttr(spec.family)}"/>`)
  if (spec.complexFamily) children.push(`<a:cs typeface="${xmlAttr(spec.complexFamily)}"/>`)

  const attributeText = attrs.map(([k, v]) => ` ${k}="${v}"`).join('')
  if (children.length === 0) return `<a:rPr${attributeText}/>`
  return `<a:rPr${attributeText}>${children.join('')}</a:rPr>`
}

/** The run-level property names a paragraph or text body can supply. */
export const RUN_PROPERTY_KEYS = [
  'family', 'complexFamily', 'size', 'bold', 'italic', 'underline', 'color',
  'transparency', 'strike', 'caps', 'spacing', 'baseline', 'highlight', 'lang'
]

/**
 * Build a single `<a:r>`.
 * @param {string|object} run - a plain string, or `{ text, ...runProperties }`.
 * @param {object} [defaults] - body-level properties merged underneath.
 * @returns {string}
 */
export function buildRun(run, defaults = {}) {
  const spec = typeof run === 'string' || typeof run === 'number' ? { text: String(run) } : { ...run }
  const text = spec.text === undefined || spec.text === null ? '' : String(spec.text)
  const props = { ...defaults }
  for (const key of RUN_PROPERTY_KEYS) {
    if (spec[key] !== undefined) props[key] = spec[key]
  }
  return `<a:r>${buildRunProperties(props)}<a:t>${escapeXml(text)}</a:t></a:r>`
}

/** The paragraph-level property names. */
const PARAGRAPH_PROPERTY_KEYS = [
  'alignment', 'level', 'bullet', 'numbered', 'bulletType', 'bulletCharacter',
  'lineSpacing', 'spaceBefore', 'spaceAfter', 'marginLeft', 'indent'
]

/**
 * Build `<a:pPr>` for one paragraph.
 *
 * Space before/after accept a DrawingML length (`12pt`) or a bare number of
 * points, which is what the option name suggests.
 *
 * @param {object} spec
 * @returns {string}
 */
export function buildParagraphProperties(spec = {}) {
  const attrs = []
  if (spec.marginLeft !== undefined) attrs.push(['marL', String(toEmu(spec.marginLeft, 'marginLeft'))])
  if (spec.indent !== undefined) attrs.push(['indent', String(toEmu(spec.indent, 'indent'))])
  if (spec.alignment !== undefined) attrs.push(['algn', normalizeAlignment(spec.alignment)])
  if (spec.level !== undefined && spec.level !== null) {
    const level = Number(spec.level)
    if (!Number.isInteger(level) || level < 0 || level > 8) {
      throw new Error(`paragraph level must be an integer 0..8, got "${spec.level}"`)
    }
    attrs.push(['lvl', String(level)])
  }

  const children = []

  if (spec.lineSpacing !== undefined && spec.lineSpacing !== null) {
    const spacing = Number(spec.lineSpacing)
    if (!Number.isFinite(spacing) || spacing <= 0) {
      throw new Error(`lineSpacing must be greater than zero, got "${spec.lineSpacing}"`)
    }
    // DrawingML expresses proportional line spacing only: 100000 is single
    // spacing. The "exactly N pt" form is a separate child this API does not
    // expose, because a business caller means the proportional one.
    children.push(`<a:lnSpc><a:spcPct val="${Math.round(spacing * 100000)}"/></a:lnSpc>`)
  }
  if (spec.spaceBefore !== undefined && spec.spaceBefore !== null) {
    children.push(`<a:spcBef><a:spcPts val="${pointsToHundredths(spec.spaceBefore, 'spaceBefore')}"/></a:spcBef>`)
  }
  if (spec.spaceAfter !== undefined && spec.spaceAfter !== null) {
    children.push(`<a:spcAft><a:spcPts val="${pointsToHundredths(spec.spaceAfter, 'spaceAfter')}"/></a:spcAft>`)
  }

  const bulletRequested = spec.bullet !== undefined || spec.numbered !== undefined
    || spec.bulletCharacter !== undefined
  if (bulletRequested) {
    const explicitNone = spec.bullet === false || spec.bullet === 'none'
      || spec.numbered === false
    if (explicitNone) {
      children.push('<a:buNone/>')
    } else if (spec.numbered || spec.bullet === 'numbered') {
      const type = typeof spec.numbered === 'string'
        ? spec.numbered
        : (spec.bulletType || 'arabicPeriod')
      children.push(`<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${xmlAttr(type)}" startAt="1"/>`)
    } else {
      const character = typeof spec.bullet === 'string' && spec.bullet !== 'true'
        ? spec.bullet
        : (spec.bulletCharacter || '•')
      children.push(
        `<a:buFont typeface="${xmlAttr(bulletFont(character))}"/>`
        + `<a:buChar char="${xmlAttr(character)}"/>`
      )
    }
  }

  const attributeText = attrs.map(([k, v]) => ` ${k}="${v}"`).join('')
  if (children.length === 0) return `<a:pPr${attributeText}/>`
  return `<a:pPr${attributeText}>${children.join('')}</a:pPr>`
}

/**
 * Space before/after to the 1/100 pt units `<a:spcPts>` uses.
 *
 * A bare number is read as points, because "spaceAfter: 12" means 12 pt to
 * everybody who is not holding a DrawingML spec.
 */
function pointsToHundredths(value, field) {
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value < 0) {
      throw new Error(`${field} must be a non-negative number of points, got "${value}"`)
    }
    return Math.round(value * 100)
  }
  const emu = toEmu(value, field)
  return Math.round(emu / 127)
}

/** A font that actually contains a given bullet glyph. */
function bulletFont(character) {
  if (['•', '–', '—', '·', '◦', '▪'].includes(character)) return 'Arial'
  return '+mn-lt'
}

/**
 * Build the `<a:p>` sequence for a text body.
 *
 * @param {Array<string|object>} paragraphs - each entry is a string or
 *   `{ text | runs, ...paragraphProperties, ...runProperties }`.
 * @param {object} [defaults] - body-level run properties.
 * @returns {string}
 */
export function buildParagraphs(paragraphs = [], defaults = {}) {
  const out = []
  for (const item of paragraphs) {
    if (item === null || item === undefined) continue
    if (typeof item === 'string' || typeof item === 'number') {
      out.push(`<a:p>${buildRun(String(item), defaults)}</a:p>`)
      continue
    }

    // Run defaults inherit the body defaults, then the paragraph's own
    // run-level properties, which is the order a caller expects them read.
    const runDefaults = { ...defaults }
    for (const key of RUN_PROPERTY_KEYS) {
      if (item[key] !== undefined) runDefaults[key] = item[key]
    }

    let runs
    if (Array.isArray(item.runs)) {
      runs = item.runs.map((run) => buildRun(run, runDefaults)).join('')
    } else {
      const text = item.text === undefined || item.text === null ? '' : String(item.text)
      runs = text.split('\n')
        .map((line, i) => (i === 0 ? buildRun(line, runDefaults) : `<a:br>${buildRun(line, runDefaults)}</a:br>`))
        .join('')
    }

    out.push(`<a:p>${buildParagraphProperties(item)}${runs}</a:p>`)
  }
  if (out.length === 0) out.push('<a:p/>')
  return out.join('')
}

/**
 * Build a complete `<p:txBody>`.
 *
 * @param {Array<string|object>} paragraphs
 * @param {object} [spec] - verticalAnchor, wrap, insets, autofit, plus run
 *   defaults applied to every run.
 * @returns {string}
 */
export function buildTextBody(paragraphs = [], spec = {}) {
  const attrs = []
  if (spec.verticalAnchor !== undefined) attrs.push(` anchor="${normalizeAnchor(spec.verticalAnchor)}"`)
  if (spec.wrap !== undefined) attrs.push(` wrap="${spec.wrap ? 'square' : 'none'}"`)
  const insets = [
    ['insetLeft', 'lIns'], ['insetTop', 'tIns'], ['insetRight', 'rIns'], ['insetBottom', 'bIns']
  ]
  for (const [key, attr] of insets) {
    if (spec[key] !== undefined) attrs.push(` ${attr}="${toEmu(spec[key], key)}"`)
  }

  let autofit = '<a:noAutofit/>'
  if (spec.autofit === 'shrink' || spec.autofit === 'normAutofit') autofit = '<a:normAutofit/>'
  else if (spec.autofit === 'resize' || spec.autofit === 'spAutoFit') autofit = '<a:spAutoFit/>'

  const defaults = {}
  // Run defaults may be spelled at the top level or grouped under `font`, which
  // is how the object API exposes them. Both are honoured, `font` winning, so a
  // caller never has to know which shape of options the builder prefers.
  for (const key of RUN_PROPERTY_KEYS) {
    if (spec[key] !== undefined) defaults[key] = spec[key]
  }
  if (spec.font && typeof spec.font === 'object') {
    for (const key of RUN_PROPERTY_KEYS) {
      if (spec.font[key] !== undefined) defaults[key] = spec.font[key]
    }
  }

  return `<p:txBody><a:bodyPr${attrs.join('')}>${autofit}</a:bodyPr>`
    + `<a:lstStyle/>${buildParagraphs(paragraphs, defaults)}</p:txBody>`
}

// ------------------------------------------------------------------ geometry

/** Rotation in degrees to the 1/60000 degree units `rot` uses. */
export function rotationToUnits(degrees) {
  const value = Number(degrees)
  if (!Number.isFinite(value)) throw new Error(`rotation must be a number of degrees, got "${degrees}"`)
  return String(Math.round(value * 60000))
}

/** The 1/60000 degree units of `rot` back to degrees. */
export function rotationToDegrees(units) {
  if (units === null || units === undefined || units === '') return 0
  const value = Number(units)
  if (!Number.isFinite(value)) return 0
  return Math.round((value / 60000) * 100) / 100
}

/** Line width in EMU. A small bare number is read as points. */
export function lineWidthToEmu(width) {
  const value = Number(width)
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`line width must be a positive length, got "${width}"`)
  }
  if (value < 100) return Math.round(value * 12700)
  return toEmu(width, 'lineWidth')
}

/**
 * The fill element for a shape, or null when the caller set no fill.
 *
 * Transparency is read from the `fill` descriptor or the explicit
 * `fillTransparency`, never from the run-level `transparency`: a bare
 * `transparency` belongs to the text, and a half-transparent text colour must
 * not silently make the shape it sits on see-through too.
 */
function resolveFill(spec) {
  const descriptor = spec.fill && typeof spec.fill === 'object' ? spec.fill : null
  if (spec.noFill === true || spec.fill === null || spec.fill === 'none') return '<a:noFill/>'
  const colour = descriptor ? descriptor.color : (spec.fill || spec.fillColor)
  if (!colour) return null
  const transparency = descriptor && descriptor.transparency !== undefined
    ? descriptor.transparency
    : spec.fillTransparency
  return `<a:solidFill>${srgbClrElement(colour, transparency)}</a:solidFill>`
}

/**
 * The `<a:ln>` element for a shape, or null when the caller set no stroke.
 *
 * A stroke may be spelled either as a `line` descriptor or as flat
 * `lineColor`/`lineWidth`/`lineDash` options; both reach the same element.
 */
function resolveLine(spec) {
  const descriptor = spec.line && typeof spec.line === 'object' ? spec.line : null
  if (spec.noLine === true || spec.line === null || spec.line === 'none') {
    return '<a:ln><a:noFill/></a:ln>'
  }

  const colour = descriptor ? descriptor.color : (spec.line === undefined ? spec.lineColor : spec.line)
  const width = descriptor && descriptor.width !== undefined ? descriptor.width : spec.lineWidth
  const transparency = descriptor && descriptor.transparency !== undefined
    ? descriptor.transparency
    : spec.lineTransparency
  const dash = (descriptor && descriptor.dash) || spec.lineDash || spec.lineStyle || spec.dash
  const arrows = resolveArrows(
    descriptor && descriptor.arrows !== undefined ? descriptor.arrows : spec.arrows
  )

  if (!colour && width === undefined && !dash && !arrows.head && !arrows.tail) return null

  const attrs = []
  if (width !== undefined && width !== null) attrs.push(` w="${lineWidthToEmu(width)}"`)

  const children = []
  if (colour) {
    children.push(`<a:solidFill>${srgbClrElement(colour, transparency)}</a:solidFill>`)
  }
  if (dash) children.push(`<a:prstDash val="${xmlAttr(dash)}"/>`)
  if (arrows.head) children.push(`<a:headEnd type="${arrows.head}" w="med" len="med"/>`)
  if (arrows.tail) children.push(`<a:tailEnd type="${arrows.tail}" w="med" len="med"/>`)

  return `<a:ln${attrs.join('')}>${children.join('')}</a:ln>`
}

/**
 * Build `<p:spPr>`.
 *
 * @param {object} spec - x, y, width, height, rotation, flipHorizontal,
 *   flipVertical, preset, customPath, fill, transparency, noFill, line,
 *   lineWidth, lineTransparency, dash, noLine, arrows, shadow.
 * @returns {string}
 */
export function buildShapeProperties(spec = {}) {
  const parts = []

  const hasGeometry = ['x', 'y', 'width', 'height'].some((k) => spec[k] !== undefined)
  if (hasGeometry) {
    const xfrmAttrs = []
    if (spec.rotation) xfrmAttrs.push(` rot="${rotationToUnits(spec.rotation)}"`)
    if (spec.flipHorizontal) xfrmAttrs.push(' flipH="1"')
    if (spec.flipVertical) xfrmAttrs.push(' flipV="1"')
    parts.push(
      `<a:xfrm${xfrmAttrs.join('')}>`
      + `<a:off x="${spec.x === undefined ? 0 : toEmu(spec.x, 'x')}" y="${spec.y === undefined ? 0 : toEmu(spec.y, 'y')}"/>`
      + `<a:ext cx="${spec.width === undefined ? 0 : toEmu(spec.width, 'width')}" cy="${spec.height === undefined ? 0 : toEmu(spec.height, 'height')}"/>`
      + '</a:xfrm>'
    )
  }

  parts.push(spec.customPath || `<a:prstGeom prst="${xmlAttr(resolvePreset(spec.preset || 'rect'))}"><a:avLst/></a:prstGeom>`)

  const fill = resolveFill(spec)
  if (fill) parts.push(fill)
  const line = resolveLine(spec)
  if (line) parts.push(line)
  if (spec.shadow === false) parts.push('<a:effectLst/>')

  return `<p:spPr>${parts.join('')}</p:spPr>`
}

/**
 * The `<p:style>` block R7 writes for a newly drawn shape.
 *
 * It is what makes a shape follow the theme (accent1 fill, accent1 line,
 * themed text colour), and it is what lets a later theme change recolour the
 * whole deck coherently.
 */
export function defaultShapeStyle() {
  return '<p:style>'
    + '<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>'
    + '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>'
    + '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
    + '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'
    + '</p:style>'
}

/**
 * Build a `<p:sp>` (or a `<p:cxnSp>` for a connector).
 *
 * A `<p:cxnSp>` has no text body, so a caller that asks for text on a
 * connector gets a plain shape instead — a connector with a caption is two
 * objects, and pretending otherwise produces a file R7 opens but cannot edit.
 *
 * @param {object} spec
 * @returns {string}
 */
export function buildShape(spec = {}) {
  const id = Number(spec.id || 2)
  const name = spec.name || `Фигура ${id}`
  const tag = spec.shapeTag || (spec.connector ? 'p:cxnSp' : 'p:sp')

  const nvPr = spec.placeholder
    ? `<p:nvPr><p:ph${placeholderAttributes(spec.placeholder)}/></p:nvPr>`
    : '<p:nvPr/>'

  const cNvPr = spec.description
    ? `<p:cNvPr id="${id}" name="${xmlAttr(name)}" descr="${xmlAttr(spec.description)}"/>`
    : `<p:cNvPr id="${id}" name="${xmlAttr(name)}"/>`

  const textBody = spec.textBody !== undefined
    ? spec.textBody
    : (spec.paragraphs ? buildTextBody(spec.paragraphs, spec) : '')

  const style = spec.style === null || spec.style === false ? '' : (spec.style || defaultShapeStyle())

  if (tag === 'p:cxnSp') {
    return '<p:cxnSp><p:nvCxnSpPr>'
      + cNvPr
      + '<p:cNvCxnSpPr/>'
      + nvPr
      + '</p:nvCxnSpPr>'
      + buildShapeProperties(spec)
      + (style ? style : '')
      + '</p:cxnSp>'
  }

  return '<p:sp><p:nvSpPr>'
    + cNvPr
    + '<p:cNvSpPr/>'
    + nvPr
    + '</p:nvSpPr>'
    + buildShapeProperties(spec)
    + textBody
    + (style || '')
    + '</p:sp>'
}

/** Serialise a placeholder descriptor into `p:ph` attributes. */
export function placeholderAttributes(placeholder) {
  const attrs = []
  if (placeholder.type) attrs.push(` type="${xmlAttr(placeholder.type)}"`)
  if (placeholder.idx !== undefined && placeholder.idx !== null) attrs.push(` idx="${xmlAttr(placeholder.idx)}"`)
  if (placeholder.orient) attrs.push(` orient="${xmlAttr(placeholder.orient)}"`)
  if (placeholder.size) attrs.push(` sz="${xmlAttr(placeholder.size)}"`)
  return attrs.join('')
}

/**
 * Build a `<p:pic>`.
 *
 * @param {object} spec - id, name, description, relId, x, y, width, height,
 *   rotation, flipHorizontal, flipVertical, hyperlinkRelId, srcRect.
 * @returns {string}
 */
export function buildPicture(spec = {}) {
  const id = Number(spec.id || 2)
  const name = spec.name || `Изображение ${id}`
  const descr = spec.description ? ` descr="${xmlAttr(spec.description)}"` : ''

  const xfrmAttrs = []
  if (spec.rotation) xfrmAttrs.push(` rot="${rotationToUnits(spec.rotation)}"`)
  if (spec.flipHorizontal) xfrmAttrs.push(' flipH="1"')
  if (spec.flipVertical) xfrmAttrs.push(' flipV="1"')

  const spPr = `<p:spPr><a:xfrm${xfrmAttrs.join('')}>`
    + `<a:off x="${toEmu(spec.x ?? 0, 'x')}" y="${toEmu(spec.y ?? 0, 'y')}"/>`
    + `<a:ext cx="${toEmu(spec.width ?? 0, 'width')}" cy="${toEmu(spec.height ?? 0, 'height')}"/>`
    + '</a:xfrm>'
    + `<a:prstGeom prst="${xmlAttr(spec.preset || 'rect')}"><a:avLst/></a:prstGeom>`
    + '</p:spPr>'

  const hyperlink = spec.hyperlinkRelId
    ? `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${xmlAttr(spec.hyperlinkRelId)}"/>`
    : ''

  const blip = '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    + ` r:embed="${xmlAttr(spec.relId)}">${hyperlink}</a:blip>`

  const srcRect = spec.srcRect
    ? `<a:srcRect${['l', 't', 'r', 'b'].filter((k) => spec.srcRect[k] !== undefined)
      .map((k) => ` ${k}="${xmlAttr(spec.srcRect[k])}"`).join('')}/>`
    : ''

  return '<p:pic><p:nvPicPr>'
    + `<p:cNvPr id="${id}" name="${xmlAttr(name)}"${descr}/>`
    + '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
    + '<p:nvPr/>'
    + '</p:nvPicPr>'
    + `<p:blipFill>${blip}${srcRect}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
    + spPr
    + '</p:pic>'
}

// ------------------------------------------------------- patching live markup

/**
 * Apply run properties to every run of a text body.
 *
 * A run that inherits its look from the layout gains an explicit `<a:rPr>`,
 * which is what "make this text bold" has to mean on a placeholder.
 *
 * @param {string} xml - markup containing a `<p:txBody>`.
 * @param {object} font - run properties, as {@link buildRunProperties} takes.
 * @returns {string}
 */
export function patchTextBodyFont(xml, font) {
  return patchTextBodyParagraphs(xml, (paragraphXml) => patchParagraphRuns(paragraphXml, font))
}

/**
 * Apply paragraph properties to every paragraph of a text body.
 * @param {string} xml
 * @param {object} spec
 * @returns {string}
 */
export function patchTextBodyParagraphProperties(xml, spec) {
  return patchTextBodyParagraphs(xml, (paragraphXml) => patchParagraphProperties(paragraphXml, spec))
}

/**
 * Set an attribute on the `<a:bodyPr>` of a text body — the element that owns
 * vertical anchoring, insets and wrapping.
 *
 * @param {string} xml
 * @param {string} name - e.g. `anchor`, `wrap`, `lIns`.
 * @param {string|null} value - null removes the attribute.
 * @returns {string}
 */
export function patchTextBodyAttribute(xml, name, value) {
  const bodyPr = firstElement(xml, 'a:bodyPr')
  if (!bodyPr) return xml
  const updated = withAttribute(bodyPr, name, value)
  return xml.replace(bodyPr, updated)
}

/**
 * Make a text body fit the placeholder it lives in instead of growing it.
 *
 * A slide-level placeholder inherits its box from the layout, so it must never
 * carry a transform of its own. The price of an inherited box is that a long
 * line has nowhere to go: without an autofit rule the renderer draws the text
 * past the bottom of the box and on top of whatever follows. `<a:normAutofit/>`
 * tells the renderer to shrink the text to the box, which is what every editor
 * writes for a title.
 *
 * @param {string} xml - a shape, or markup containing one `<p:txBody>`.
 * @returns {string}
 */
export function ensureNormAutofit(xml) {
  const bodyPr = firstElement(xml, 'a:bodyPr')
  if (!bodyPr) return xml
  let updated = removeAll(bodyPr, 'a:noAutofit')
  updated = removeAll(updated, 'a:spAutoFit')
  if (!firstElement(updated, 'a:normAutofit')) {
    updated = setChild(updated, 'a:normAutofit', '<a:normAutofit/>', BODYPR_ORDER)
  }
  if (updated === bodyPr) return xml
  return xml.replace(bodyPr, updated)
}

/** Map every `<a:p>` inside the `<p:txBody>` of `xml` through `fn`. */
function patchTextBodyParagraphs(xml, fn) {
  const body = findElement(xml, 'p:txBody')
  if (!body) return xml

  const paragraphs = extractElements(body.innerXml, 'a:p')
  let newInner = body.innerXml
  let offset = 0

  for (const paragraph of paragraphs) {
    const at = paragraph.index + offset
    // A paragraph inside a text body cannot contain another `a:p`, so mapping
    // each one in document order is exact.
    const updated = fn(paragraph.outerXml)
    newInner = newInner.slice(0, at) + updated + newInner.slice(at + paragraph.outerXml.length)
    offset += updated.length - paragraph.outerXml.length
  }

  const updatedBody = body.outerXml.replace(body.innerXml, newInner)
  return xml.slice(0, body.start) + updatedBody + xml.slice(body.end)
}

/**
 * Apply run properties to one paragraph of a text body.
 *
 * @param {string} xml
 * @param {number} paragraphIndex
 * @param {object} font
 * @returns {string}
 */
export function patchParagraphFont(xml, paragraphIndex, font) {
  return patchOneTextBodyParagraph(xml, paragraphIndex, (paragraphXml) => patchParagraphRuns(paragraphXml, font))
}

/** Map exactly one `<a:p>` of the text body through `fn`. */
function patchOneTextBodyParagraph(xml, paragraphIndex, fn) {
  const body = findElement(xml, 'p:txBody')
  if (!body) return xml
  const paragraphs = extractElements(body.innerXml, 'a:p')
  const target = paragraphs[paragraphIndex]
  if (!target) {
    throw new Error(
      `Paragraph index ${paragraphIndex} is out of range: the object has ${paragraphs.length} paragraph(s)`
    )
  }
  const updatedParagraph = fn(target.outerXml)
  const newInner = body.innerXml.slice(0, target.index)
    + updatedParagraph
    + body.innerXml.slice(target.index + target.outerXml.length)
  const updatedBody = body.outerXml.replace(body.innerXml, newInner)
  return xml.slice(0, body.start) + updatedBody + xml.slice(body.end)
}

/**
 * Apply both paragraph properties and run properties to a single paragraph.
 * @param {string} xml
 * @param {number} paragraphIndex
 * @param {object} spec
 * @param {object|null} font
 * @returns {string}
 */
export function patchOneParagraph(xml, paragraphIndex, spec, font = null) {
  return patchOneTextBodyParagraph(xml, paragraphIndex, (paragraphXml) => {
    const withProperties = patchParagraphProperties(paragraphXml, spec)
    return font ? patchParagraphRuns(withProperties, font) : withProperties
  })
}

/** Patch one `<a:p>`: every run gains the merged properties. */
function patchParagraphRuns(paragraphXml, font) {
  let out = paragraphXml
  let offset = 0

  for (const run of extractElements(paragraphXml, 'a:r')) {
    const at = run.index + offset
    const updated = patchRunFont(run.outerXml, font)
    out = out.slice(0, at) + updated + out.slice(at + run.outerXml.length)
    offset += updated.length - run.outerXml.length
  }

  // A layout placeholder holds only an end-paragraph mark until someone types.
  // Without this the restyled slide reads back with no font at all, and the
  // next thing typed reverts to the layout's look.
  if (extractElements(paragraphXml, 'a:r').length === 0) {
    const endPara = firstElement(out, 'a:endParaRPr')
    if (endPara) {
      const merged = mergeRunProperties(endPara, font)
      const at = out.indexOf(endPara)
      out = out.slice(0, at) + merged + out.slice(at + endPara.length)
    } else {
      out = out.replace('</a:p>', `${buildRunProperties(font)}</a:p>`)
    }
  }
  return out
}

/** Replace (or insert) `<a:rPr>` inside a single `<a:r>`. */
function patchRunFont(runXml, font) {
  const existing = firstElement(runXml, 'a:rPr')
  if (existing) {
    const merged = mergeRunProperties(existing, font)
    const at = runXml.indexOf(existing)
    return runXml.slice(0, at) + merged + runXml.slice(at + existing.length)
  }
  const built = buildRunProperties({ ...font, dirty: 0 })
  const tElement = firstElement(runXml, 'a:t')
  if (tElement) {
    const at = runXml.indexOf(tElement)
    return runXml.slice(0, at) + built + runXml.slice(at)
  }
  return runXml.replace(/(<\/a:r>|<a:r\/>)$/, `${built}$1`)
}

/**
 * Merge requested run properties into an existing `<a:rPr>` or
 * `<a:endParaRPr>`.
 *
 * Attributes are replaced in place and only the children the caller named are
 * touched, so kerning, language, hyperlinks and unknown extensions survive.
 *
 * @param {string} rPrXml
 * @param {object} font
 * @returns {string}
 */
export function mergeRunProperties(rPrXml, font) {
  let out = rPrXml

  if (font.size !== undefined && font.size !== null) {
    out = withAttribute(out, 'sz', String(Math.round(Number(font.size) * 100)))
  }
  if (font.bold !== undefined) out = withAttribute(out, 'b', font.bold ? '1' : '0')
  if (font.italic !== undefined) out = withAttribute(out, 'i', font.italic ? '1' : '0')
  if (font.underline !== undefined) out = withAttribute(out, 'u', normalizeUnderline(font.underline))
  if (font.strike !== undefined) out = withAttribute(out, 'strike', font.strike ? 'sngStrike' : 'noStrike')
  if (font.spacing !== undefined && font.spacing !== null) {
    out = withAttribute(out, 'spc', String(Math.round(Number(font.spacing) * 100)))
  }
  if (font.baseline !== undefined && font.baseline !== null) {
    out = withAttribute(out, 'baseline', String(Math.round(Number(font.baseline) * 1000)))
  }
  if (font.caps !== undefined) {
    const caps = String(font.caps).toLowerCase()
    if (caps === 'small') out = withAttribute(out, 'cap', 'small')
    else if (caps === 'all' || caps === 'uppercase') out = withAttribute(out, 'cap', 'all')
    else out = withAttribute(out, 'cap', 'none')
  }

  if (font.color !== undefined && font.color !== null) {
    out = setChild(out, 'a:solidFill', `<a:solidFill>${srgbClrElement(font.color, font.transparency)}</a:solidFill>`, RPR_ORDER)
  }
  if (font.solidFillXml) {
    out = setChild(out, 'a:solidFill', font.solidFillXml, RPR_ORDER)
  }
  if (font.highlight) {
    out = setChild(out, 'a:highlight', `<a:highlight>${srgbClrElement(font.highlight)}</a:highlight>`, RPR_ORDER)
  }
  if (font.family) {
    out = setChild(out, 'a:latin', `<a:latin typeface="${xmlAttr(font.family)}"/>`, RPR_ORDER)
    // Latin, complex-script and East-Asian are separate slots. Naming only
    // `a:latin` leaves Cyrillic to the theme font in R7, which is exactly the
    // bug a caller reports as "the font did not change".
    if (font.complexFamily !== false) {
      out = setChild(out, 'a:cs', `<a:cs typeface="${xmlAttr(font.family)}"/>`, RPR_ORDER)
    }
  }
  return out
}

/** Merge paragraph properties into one `<a:p>`. */
export function patchParagraphProperties(paragraphXml, spec) {
  const existing = firstElement(paragraphXml, 'a:pPr')
  let pPr = existing || '<a:pPr/>'

  if (spec.alignment !== undefined && spec.alignment !== null) {
    pPr = withAttribute(pPr, 'algn', normalizeAlignment(spec.alignment))
  }
  if (spec.level !== undefined && spec.level !== null) {
    pPr = withAttribute(pPr, 'lvl', String(Number(spec.level)))
  }
  if (spec.marginLeft !== undefined && spec.marginLeft !== null) {
    pPr = withAttribute(pPr, 'marL', String(toEmu(spec.marginLeft, 'marginLeft')))
  }
  if (spec.indent !== undefined && spec.indent !== null) {
    pPr = withAttribute(pPr, 'indent', String(toEmu(spec.indent, 'indent')))
  }

  if (spec.lineSpacing !== undefined && spec.lineSpacing !== null) {
    const percent = Math.round(Number(spec.lineSpacing) * 100000)
    pPr = setChild(pPr, 'a:lnSpc', `<a:lnSpc><a:spcPct val="${percent}"/></a:lnSpc>`, PPR_ORDER)
  }
  if (spec.spaceBefore !== undefined && spec.spaceBefore !== null) {
    pPr = setChild(pPr, 'a:spcBef', `<a:spcBef><a:spcPts val="${pointsToHundredths(spec.spaceBefore, 'spaceBefore')}"/></a:spcBef>`, PPR_ORDER)
  }
  if (spec.spaceAfter !== undefined && spec.spaceAfter !== null) {
    pPr = setChild(pPr, 'a:spcAft', `<a:spcAft><a:spcPts val="${pointsToHundredths(spec.spaceAfter, 'spaceAfter')}"/></a:spcAft>`, PPR_ORDER)
  }

  if (spec.bullet !== undefined || spec.numbered !== undefined || spec.bulletCharacter !== undefined) {
    for (const tag of ['a:buNone', 'a:buChar', 'a:buAutoNum', 'a:buFont', 'a:buClr', 'a:buSzPct', 'a:buSzPts']) {
      pPr = removeAll(pPr, tag)
    }
    const none = spec.bullet === false || spec.bullet === 'none' || spec.numbered === false
    if (none) {
      pPr = setChild(pPr, 'a:buNone', '<a:buNone/>', PPR_ORDER)
    } else if (spec.numbered && spec.numbered !== 'none') {
      const type = typeof spec.numbered === 'string' ? spec.numbered : 'arabicPeriod'
      pPr = setChild(pPr, 'a:buFont', '<a:buFont typeface="+mj-lt"/>', PPR_ORDER)
      pPr = setChild(pPr, 'a:buAutoNum', `<a:buAutoNum type="${xmlAttr(type)}" startAt="1"/>`, PPR_ORDER)
    } else {
      const character = typeof spec.bullet === 'string' && spec.bullet !== 'true'
        ? spec.bullet
        : (spec.bulletCharacter || '•')
      pPr = setChild(pPr, 'a:buFont', `<a:buFont typeface="${xmlAttr(bulletFont(character))}"/>`, PPR_ORDER)
      pPr = setChild(pPr, 'a:buChar', `<a:buChar char="${xmlAttr(character)}"/>`, PPR_ORDER)
    }
  }

  if (existing) {
    const at = paragraphXml.indexOf(existing)
    return paragraphXml.slice(0, at) + pPr + paragraphXml.slice(at + existing.length)
  }
  return paragraphXml.replace(/(<a:p(?:\s[^>]*)?>)/, `$1${pPr}`)
}

/**
 * Patch shape geometry and paint directly in a `<p:spPr>`.
 *
 * The `<p:spPr>` is rebuilt only for the properties the caller named: any
 * element this does not know about stays byte-identical in place.
 *
 * @param {string} shapeXml - a `<p:sp>`, `<p:pic>` or `<p:cxnSp>`.
 * @param {object} spec
 * @returns {string}
 */
export function patchShapeProperties(shapeXml, spec) {
  const spPr = findElement(shapeXml, 'p:spPr')
  if (!spPr) return shapeXml
  let pr = spPr.outerXml
  const xfrm = spec.xfrm || spec

  // --- geometry --------------------------------------------------------
  const touchesGeometry = ['x', 'y', 'width', 'height', 'rotation', 'flipHorizontal', 'flipVertical']
    .some((k) => xfrm[k] !== undefined)
  if (touchesGeometry) {
    const off = firstElement(pr, 'a:off')
    const ext = firstElement(pr, 'a:ext')
    const currentX = off ? Number(requireAttr(off, 'x')) : 0
    const currentY = off ? Number(requireAttr(off, 'y')) : 0
    const currentCx = ext ? Number(requireAttr(ext, 'cx')) : 0
    const currentCy = ext ? Number(requireAttr(ext, 'cy')) : 0
    const currentXfrm = firstElement(pr, 'a:xfrm')

    const attrs = []
    const rot = xfrm.rotation !== undefined
      ? xfrm.rotation
      : (currentXfrm ? Number(attrOf(currentXfrm, 'rot') || 0) / 60000 : 0)
    if (rot) attrs.push(` rot="${rotationToUnits(rot)}"`)
    if (xfrm.flipHorizontal || (!xfrm.flipHorizontal && currentXfrm && /flipH=/.test(currentXfrm))) {
      attrs.push(' flipH="1"')
    }
    if (xfrm.flipVertical || (!xfrm.flipVertical && currentXfrm && /flipV=/.test(currentXfrm))) {
      attrs.push(' flipV="1"')
    }

    const built = `<a:xfrm${attrs.join('')}>`
      + `<a:off x="${xfrm.x === undefined ? currentX : toEmu(xfrm.x, 'x')}" y="${xfrm.y === undefined ? currentY : toEmu(xfrm.y, 'y')}"/>`
      + `<a:ext cx="${xfrm.width === undefined ? currentCx : toEmu(xfrm.width, 'width')}" cy="${xfrm.height === undefined ? currentCy : toEmu(xfrm.height, 'height')}"/>`
      + '</a:xfrm>'
    pr = setChild(pr, 'a:xfrm', built, SPPR_ORDER)
  }

  // --- fill ------------------------------------------------------------
  const touchesFill = spec.fill !== undefined || spec.fillColor !== undefined
    || spec.noFill !== undefined || spec.fillTransparency !== undefined
  if (touchesFill) {
    const fill = resolveFill(spec)
    if (fill) {
      // Every fill kind is exclusive, so the existing one is dropped whole:
      // a leftover `<a:gradFill>` next to a new `<a:solidFill>` is a package
      // R7 renders arbitrarily.
      for (const tag of ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill']) {
        pr = removeAll(pr, tag)
      }
      const tag = fill.startsWith('<a:noFill') ? 'a:noFill' : 'a:solidFill'
      pr = setChild(pr, tag, fill, SPPR_ORDER)
    }
  }

  // --- stroke ----------------------------------------------------------
  const touchesLine = spec.line !== undefined || spec.lineColor !== undefined
    || spec.lineWidth !== undefined || spec.lineTransparency !== undefined
    || spec.lineStyle !== undefined || spec.lineDash !== undefined
    || spec.dash !== undefined || spec.noLine !== undefined || spec.arrows !== undefined
  if (touchesLine) {
    const line = resolveLine(spec) || '<a:ln/>'
    pr = setChild(pr, 'a:ln', line, SPPR_ORDER)
  }

  return shapeXml.replace(spPr.outerXml, pr)
}

function attrOf(tagXml, name) {
  const match = tagXml.match(new RegExp(`\\s${name}="([^"]*)"`))
  return match ? match[1] : null
}

function requireAttr(tagXml, name) {
  const value = attrOf(tagXml, name)
  if (value === null) throw new Error(`Missing attribute ${name} in <${tagXml.slice(0, 20)}…>`)
  return value
}
