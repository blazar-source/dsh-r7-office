import { extractElements, getAttribute } from '../shared/xml.js'

/**
 * Shared plumbing for the PPTX engine.
 *
 * Everything in here is pure: measurement conversion, colour conversion, tag
 * surgery on XML strings, and discovery of the presentation's part graph.
 * The engine itself only orchestrates, which keeps the tricky parts (schema
 * child order, placeholder inheritance) testable in isolation.
 */

// ---------------------------------------------------------------- measurements

/** English Metric Units per inch. DrawingML geometry is expressed in EMU. */
export const EMU_PER_INCH = 914400
/** EMU per point (1/72 inch). */
export const EMU_PER_POINT = 12700
/** EMU per centimetre. */
export const EMU_PER_CM = 360000
/** EMU per screen pixel at the CSS reference density of 96 dpi. */
export const EMU_PER_PIXEL = 9525

/** Slide size used when a presentation declares none (16:9 widescreen). */
export const DEFAULT_SLIDE_SIZE = { width: 12192000, height: 6858000 }

/**
 * Convert a user-facing length into EMU.
 *
 * Accepts a plain number (already EMU, because that is what a caller reading
 * geometry back out of `readSlide` holds), or a string/number with an explicit
 * unit: `12pt`, `30px`, `2.5cm`, `1in`, `457200emu`.
 *
 * @param {number|string} value
 * @param {string} [field] - name used in the error message.
 * @returns {number}
 */
export function toEmu(value, field = 'value') {
  if (value === null || value === undefined || value === '') {
    throw new Error(`${field} is required`)
  }
  if (typeof value === 'number') {
    assertFinite(value, field)
    return Math.round(value)
  }
  const raw = String(value).trim()
  const match = raw.match(/^(-?\d+(?:\.\d+)?)\s*(emu|pt|px|cm|mm|in|")?$/i)
  if (!match) {
    throw new Error(`${field}: cannot parse length "${raw}" (use a number of EMU, or e.g. 24pt, 30px, 2cm)`)
  }
  const amount = Number(match[1])
  assertFinite(amount, field)
  const unit = (match[2] || 'emu').toLowerCase()
  switch (unit) {
    case 'emu': return Math.round(amount)
    case 'pt': return Math.round(amount * EMU_PER_POINT)
    case 'px': return Math.round(amount * EMU_PER_PIXEL)
    case 'cm': return Math.round(amount * EMU_PER_CM)
    case 'mm': return Math.round(amount * EMU_PER_CM / 10)
    case 'in':
    case '"': return Math.round(amount * EMU_PER_INCH)
    default: throw new Error(`${field}: unsupported unit "${unit}"`)
  }
}

/** EMU to points, rounded to two decimals, for the read model. */
export function emuToPoints(emu) {
  return Math.round((emu / EMU_PER_POINT) * 100) / 100
}

/** EMU to inches, rounded to four decimals, for the read model. */
export function emuToInches(emu) {
  return Math.round((emu / EMU_PER_INCH) * 10000) / 10000
}

function assertFinite(value, field) {
  if (!Number.isFinite(value)) throw new Error(`${field}: "${value}" is not a finite number`)
}

// --------------------------------------------------------------------- colour

/**
 * Normalize a colour to the ARGB hex form DrawingML stores.
 *
 * `#RRGGBB` and `RRGGBB` gain an opaque alpha; the alpha channel of an
 * `AARRGGBB` value is respected.
 *
 * @param {string} value
 * @returns {string} upper-case `AARRGGBB`
 */
export function toArgb(value) {
  if (value === null || value === undefined) throw new Error('A colour value is required')
  let hex = String(value).trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{3}$/.test(hex)) {
    hex = hex.split('').map((c) => c + c).join('')
  }
  if (!/^[0-9A-F]{6}$/.test(hex) && !/^[0-9A-F]{8}$/.test(hex)) {
    throw new Error(`Invalid colour "${value}". Use #RRGGBB or #AARRGGBB.`)
  }
  if (hex.length === 6) hex = `FF${hex}`
  return hex
}

/** `AARRGGBB` to `#RRGGBB`, the form the read model reports. */
export function argbToHex(argb) {
  if (!argb) return null
  const hex = String(argb).toUpperCase()
  if (hex.length === 8) return `#${hex.slice(2)}`
  if (hex.length === 6) return `#${hex}`
  return null
}

/** Opacity 0..1 from an `AARRGGBB` value. */
export function alphaFromArgb(argb) {
  if (!argb || String(argb).length !== 8) return 1
  return Math.round((parseInt(String(argb).slice(0, 2), 16) / 255) * 1000) / 1000
}

/**
 * Split a colour plus optional transparency into an `AARRGGBB`.
 *
 * `transparency` is the share of the colour that shows *through* the fill
 * (0 = opaque, 1 = invisible), which is what a business user means by it.
 *
 * @param {string} color
 * @param {number} [transparency] - 0..1
 * @returns {string}
 */
export function colorWithTransparency(color, transparency) {
  const argb = toArgb(color)
  if (transparency === null || transparency === undefined) return argb
  const t = Number(transparency)
  if (!Number.isFinite(t) || t < 0 || t > 1) {
    throw new Error(`transparency must be between 0 and 1, got "${transparency}"`)
  }
  const alpha = Math.round((1 - t) * 255)
  return `${alpha.toString(16).toUpperCase().padStart(2, '0')}${argb.slice(2)}`
}

/** Inverse of {@link colorWithTransparency} for a read-model result. */
export function transparencyFromArgb(argb) {
  if (!argb || String(argb).length !== 8) return 0
  return Math.round((1 - parseInt(String(argb).slice(0, 2), 16) / 255) * 1000) / 1000
}

// ------------------------------------------------------------------- XML bits

/** Escape a value used inside a double-quoted XML attribute. */
export function xmlAttr(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Locate a single element's extent in a parent document.
 * @param {string} xml
 * @param {string} tag
 * @returns {{start: number, end: number, outerXml: string, innerXml: string}|null}
 */
export function findElement(xml, tag) {
  const found = extractElements(xml, tag)
  if (found.length === 0) return null
  const el = found[0]
  return {
    start: el.index,
    end: el.index + el.outerXml.length,
    outerXml: el.outerXml,
    innerXml: el.innerXml
  }
}

/**
 * Return the outer XML of the first `<tag ...>` element, or null.
 * @param {string} xml
 * @param {string} tag
 * @returns {string|null}
 */
export function firstElement(xml, tag) {
  const el = extractElements(xml, tag)[0]
  return el ? el.outerXml : null
}

/**
 * Return the inner XML of the first `<tag ...>` element, or null.
 * @param {string} xml
 * @param {string} tag
 * @returns {string|null}
 */
export function firstInner(xml, tag) {
  const el = extractElements(xml, tag)[0]
  return el ? el.innerXml : null
}

/** Attributes of the first opening tag of `tag`, as a plain object. */
export function firstAttributes(xml, tag) {
  const el = extractElements(xml, tag)[0]
  if (!el) return {}
  const out = {}
  const re = /([A-Za-z_:][-A-Za-z0-9_:.]*)=["']([^"']*)["']/g
  let m
  while ((m = re.exec(el.outerXml)) !== null) out[m[1]] = m[2]
  return out
}

/** Replace (or insert) an attribute on an opening tag string. */
/**
 * Set or replace an attribute on the opening tag of an element.
 *
 * Accepts either a bare opening tag (`<a:pPr lvl="1">`, `<a:rPr/>`) or a whole
 * element with children. Working on the whole element matters: a regex that
 * anchors on the trailing `>` of an element whose content ends the string
 * writes the attribute after the close tag, producing
 * `<a:pPr lvl="1"><a:buChar/></a:pPr algn="r">` — well-formed enough to parse
 * for a reader that ignores it, and silently wrong for the caller.
 *
 * @param {string} elementXml
 * @param {string} name
 * @param {string|number|null} value - null removes the attribute.
 * @returns {string}
 */
export function withAttribute(elementXml, name, value) {
  const openEnd = elementXml.indexOf('>')
  if (openEnd === -1) return elementXml
  const openTag = elementXml.slice(0, openEnd + 1)
  const rest = elementXml.slice(openEnd + 1)

  if (value === null || value === undefined) return `${removeAttribute(openTag, name)}${rest}`

  const re = new RegExp(`\\s${escapeRegex(name)}=["'][^"']*["']`, 'i')
  if (re.test(openTag)) {
    return `${openTag.replace(re, ` ${name}="${xmlAttr(value)}"`)}${rest}`
  }

  // Insert before the closing bracket, keeping a trailing slash intact.
  const selfClosing = /\/>$/.test(openTag)
  const closing = selfClosing ? '/>' : '>'
  const updated = openTag.replace(new RegExp(`${escapeRegex(closing)}$`), ` ${name}="${xmlAttr(value)}"${closing}`)
  return `${updated}${rest}`
}

/** Drop an attribute from the opening tag of an element. */
export function removeAttribute(elementXml, name) {
  const openEnd = elementXml.indexOf('>')
  if (openEnd === -1) return elementXml
  const openTag = elementXml.slice(0, openEnd + 1)
  const rest = elementXml.slice(openEnd + 1)
  const re = new RegExp(`\\s${escapeRegex(name)}=["'][^"']*["']`, 'i')
  return `${openTag.replace(re, '')}${rest}`
}

/**
 * The opening tag of an element, e.g. `<p:spPr a="1">`, including `/>`.
 *
 * Never call this on a whole document: the first `>` in a document belongs to
 * the `<?xml …?>` declaration. Pass the element, as
 * `firstElement(xml, 'p:sldLayout')` returns it.
 */
export function openingTag(outerXml) {
  if (!outerXml) return ''
  const end = outerXml.indexOf('>')
  return end === -1 ? outerXml : outerXml.slice(0, end + 1)
}

/**
 * The opening tag of the *root element* of a document.
 * @param {string} xml
 * @param {string} tag - e.g. `p:sldLayout`.
 * @returns {string}
 */
export function rootOpeningTag(xml, tag) {
  return openingTag(firstElement(xml, tag) || '')
}

/**
 * Parse the element that starts at `xml[at]`, without descending into a
 * same-named child.
 *
 * A plain `extractElements` matches nested elements too, which produces real
 * corruption when the tag is its own descendant: patching `<a:solidFill>` with
 * `extractElements` finds the copy inside the `<a:latin>` that was just written
 * and splices a second fill into the middle of it.
 *
 * @param {string} xml
 * @param {number} at
 * @returns {{start: number, end: number}|null}
 */
function scanElement(xml, at) {
  const end = xml.indexOf('>', at)
  if (end === -1) return null
  if (xml[end - 1] === '/') return { start: at, end: end + 1 }

  const nameMatch = xml.slice(at).match(/^<([A-Za-z_][-A-Za-z0-9_:.]*)/)
  if (!nameMatch) return null
  const name = nameMatch[1]
  const open = new RegExp(`<${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?=[\\s/>])`, 'g')
  const close = `</${name}>`
  let depth = 1
  let cursor = end + 1

  while (depth > 0) {
    open.lastIndex = cursor
    const nextOpen = open.exec(xml)
    const nextClose = xml.indexOf(close, cursor)
    if (nextClose === -1) return null
    if (nextOpen && nextOpen.index < nextClose) {
      const openEnd = xml.indexOf('>', nextOpen.index)
      if (openEnd !== -1 && xml[openEnd - 1] === '/') {
        cursor = openEnd + 1
      } else {
        depth++
        cursor = nextOpen.index + nextOpen[0].length
      }
    } else {
      depth--
      cursor = nextClose + close.length
      if (depth === 0) return { start: at, end: cursor }
    }
  }
  return null
}

/** Every top-level occurrence of `tag` in a fragment, in document order. */
function topLevelSpans(xml, tag) {
  const open = new RegExp(`<${tag.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?=[\\s/>])`, 'g')
  const spans = []
  let cursor = 0
  let match
  while ((match = open.exec(xml)) !== null) {
    if (match.index < cursor) continue
    const span = scanElement(xml, match.index)
    if (!span) break
    spans.push(span)
    cursor = span.end
    open.lastIndex = cursor
  }
  return spans
}

/** The tag name of the first element in a fragment. */
function rootTagName(xml) {
  const match = String(xml).match(/^\s*<([A-Za-z_][-A-Za-z0-9_:.]*)/)
  return match ? match[1] : null
}

/** The span of the first top-level `tag` in a fragment. */
function firstSpan(xml, tag) {
  return topLevelSpans(xml, tag)[0] || null
}

/**
 * Remove every top-level occurrence of `tag` from a fragment.
 * @param {string} xml
 * @param {string} tag
 * @returns {string}
 */
export function removeAll(xml, tag) {
  const spans = topLevelSpans(xml, tag)
  if (spans.length === 0) return xml
  let out = ''
  let cursor = 0
  for (const span of spans) {
    out += xml.slice(cursor, span.start)
    cursor = span.end
  }
  return out + xml.slice(cursor)
}

/**
 * Replace the first top-level `childTag` inside a fragment, or insert it at its
 * schema position.
 *
 * OOXML validates children by order, so appending an `<a:latin>` after an
 * `<a:hlinkClick>` produces a package PowerPoint refuses to open. `order`
 * lists the element's known children in schema order; the new child is placed
 * before the first present child that follows it, or just before the closing
 * tag when none does.
 *
 * The fragment may be self-closing (`<a:rPr/>`), which is what properties are
 * before anything styles them.
 *
 * @param {string} xml - a single element, or the whole document.
 * @param {string} childTag - tag name of the child, e.g. `a:latin`.
 * @param {string} childXml - the element to place.
 * @param {string[]} order - schema order of the element's children.
 * @returns {string}
 */
export function setChild(xml, childTag, childXml, order) {
  const containerTag = rootTagName(xml)
  const existing = firstSpan(xml, childTag)
  const out = existing
    ? xml.slice(0, existing.start) + childXml + xml.slice(existing.end)
    : xml

  // A self-closing container (`<a:rPr/>`, `<a:solidFill/>`) has to gain a body
  // before anything can live inside it. The closing tag written here is the
  // *container's*, not the child's.
  if (!existing && containerTag && new RegExp(`<${escapeRegex(containerTag)}(?:\\s[^>]*)?/>\\s*$`).test(out)) {
    return out.replace(/\/>\s*$/, `>${childXml}</${containerTag}>`)
  }
  if (existing) return out

  const rank = order.indexOf(childTag)
  const closeAt = out.lastIndexOf('</')
  let insertAt = closeAt === -1 ? out.length : closeAt

  for (const candidate of order) {
    if (candidate === childTag) continue
    const candidateRank = order.indexOf(candidate)
    if (candidateRank !== -1 && candidateRank < rank) continue
    const span = firstSpan(out, candidate)
    if (span && span.start < insertAt) insertAt = span.start
  }

  return out.slice(0, insertAt) + childXml + out.slice(insertAt)
}

/**
 * Insert `childXml` into the first `<tag>` of `xml` at its schema position.
 * @param {string} xml
 * @param {string} tag - container element name, e.g. `p:spPr`.
 * @param {string} childTag
 * @param {string} childXml
 * @param {string[]} order
 * @returns {string}
 */
export function setOrderedChild(xml, tag, childTag, childXml, order) {
  const container = findElement(xml, tag)
  if (!container) throw new Error(`Cannot set <${childTag}>: <${tag}> not found`)
  const updated = setChild(container.outerXml, childTag, childXml, order)
  return xml.slice(0, container.start) + updated + xml.slice(container.end)
}

/** Insert `childXml` immediately before the closing tag of `tag`. */
export function appendInside(xml, tag, childXml) {
  const container = findElement(xml, tag)
  if (!container) throw new Error(`Cannot append into <${tag}>: element not found`)
  const updated = container.outerXml.replace(container.innerXml, container.innerXml + childXml)
  return xml.slice(0, container.start) + updated + xml.slice(container.end)
}

/** Remove every occurrence of `tag` from `xml`. */
export function removeElements(xml, tag) {
  const found = extractElements(xml, tag)
  let out = xml
  for (let i = found.length - 1; i >= 0; i--) {
    const el = found[i]
    out = out.slice(0, el.index) + out.slice(el.index + el.outerXml.length)
  }
  return out
}

// -------------------------------------------------------------- part graph

/** Relationship types the engine needs to recognise. */
export const REL = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
}

const CONTENT_TYPES = {
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'
}

/**
 * Resolve an OPC relationship target against the part that declares it.
 *
 * Targets are relative to the *directory of the source part*, so
 * `../slideLayouts/slideLayout1.xml` seen from `ppt/slides/slide1.xml` becomes
 * `ppt/slideLayouts/slideLayout1.xml`.
 *
 * @param {string} sourcePart - e.g. `ppt/slides/slide1.xml`
 * @param {string} target - e.g. `../slideLayouts/slideLayout1.xml`
 * @returns {string}
 */
export function resolvePartPath(sourcePart, target) {
  const clean = String(target).split('#')[0]
  if (clean.startsWith('/')) return clean.slice(1)
  const base = sourcePart.split('/').slice(0, -1)
  const segments = clean.split('/')
  const stack = [...base]
  for (const segment of segments) {
    if (!segment || segment === '.') continue
    if (segment === '..') stack.pop()
    else stack.push(segment)
  }
  return stack.join('/')
}

/** Where the relationship part for a given part lives. */
export function relsPathFor(partPath) {
  const dir = partPath.split('/').slice(0, -1).join('/')
  const file = partPath.split('/').pop()
  return `${dir ? `${dir}/` : ''}_rels/${file}.rels`
}

/**
 * All relationships declared by a part.
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} partPath
 * @returns {Array<{id: string, type: string, target: string, targetMode: string|null, partPath: string}>}
 */
export function readRelationships(zip, partPath) {
  const relsXml = zip.getText(relsPathFor(partPath))
  if (!relsXml) return []
  return extractElements(relsXml, 'Relationship').map((el) => {
    const type = getAttribute(el.outerXml, 'Type') || ''
    const target = getAttribute(el.outerXml, 'Target') || ''
    const targetMode = getAttribute(el.outerXml, 'TargetMode')
    return {
      id: getAttribute(el.outerXml, 'Id') || '',
      type,
      target,
      targetMode,
      partPath: resolvePartPath(partPath, target)
    }
  })
}

/** Unique `rIdN` not yet used in a relationships part. */
export function nextRelId(relsXml) {
  const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))
  return `rId${ids.length > 0 ? Math.max(...ids) + 1 : 1}`
}

/** The next free numeric suffix for `prefix<N><suffix>` across member names. */
export function nextPartNumber(names, regex, fallback = 1) {
  const used = []
  for (const name of names) {
    const m = name.match(regex)
    if (m) used.push(Number(m[1]))
  }
  return used.length > 0 ? Math.max(...used) + 1 : fallback
}

// ------------------------------------------------------------- deck structure

/**
 * The deck's slide order, resolved through `ppt/presentation.xml.rels`.
 *
 * The order of `<p:sldId>` in the presentation is the authoring order and has
 * nothing to do with the file name of the slide part: deleting slide 2 of 4
 * leaves `slide1.xml`, `slide3.xml`, `slide4.xml` on disk. Every operation in
 * the engine therefore addresses slides through this list, never by guessing
 * `slide<index + 1>.xml`.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @returns {Array<{index: number, sldId: number, relId: string, partPath: string, relsPath: string}>}
 */
export function slideParts(zip) {
  const presentationXml = zip.getText('ppt/presentation.xml')
  if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

  const rels = readRelationships(zip, 'ppt/presentation.xml')
  const byId = new Map(rels.map((r) => [r.id, r]))

  const entries = extractElements(presentationXml, 'p:sldId')
  const result = []
  entries.forEach((el, index) => {
    const relId = getAttribute(el.outerXml, 'r:id') || getAttribute(el.outerXml, 'relationships:id')
    const rel = relId ? byId.get(relId) : null
    if (!rel) return
    result.push({
      index,
      sldId: Number(getAttribute(el.outerXml, 'id') || 0),
      relId,
      partPath: rel.partPath,
      relsPath: relsPathFor(rel.partPath)
    })
  })
  return result
}

/**
 * The presentation's slide layouts in master order, with their display names
 * and their PowerPoint layout type.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @returns {Array<{index: number, partPath: string, name: string, type: string, placeholders: object[]}>}
 */
export function slideLayouts(zip) {
  const presentationXml = zip.getText('ppt/presentation.xml')
  if (!presentationXml) throw new Error('Invalid PPTX: ppt/presentation.xml not found')

  const layouts = []
  const masterRels = readRelationships(zip, 'ppt/presentation.xml')
    .filter((r) => r.type === REL.slideMaster)

  for (const master of masterRels) {
    const masterXml = zip.getText(master.partPath)
    // A layout whose placeholder declares no transform (R7's title layout does
    // exactly that) takes its box from the master's matching placeholder.
    const masterPlaceholders = new Map(
      layoutPlaceholders(masterXml || '').map((ph) => [`${ph.type}|${ph.idx === null ? '' : ph.idx}`, ph])
    )
    const layoutRels = readRelationships(zip, master.partPath)
      .filter((r) => r.type === REL.slideLayout)
    for (const rel of layoutRels) {
      const xml = zip.getText(rel.partPath)
      if (!xml) continue
      layouts.push({
        index: layouts.length,
        partPath: rel.partPath,
        name: layoutName(xml),
        type: getAttribute(rootOpeningTag(xml, 'p:sldLayout'), 'type') || 'custom',
        placeholders: layoutPlaceholders(xml, masterPlaceholders)
      })
    }
  }

  // A deck with no master relationship still deserves a usable answer: fall
  // back to the layout parts that physically exist, in file-name order.
  if (layouts.length === 0) {
    const names = zip.list()
      .filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n))
      .sort((a, b) => {
        const na = Number(a.match(/(\d+)\.xml$/)[1])
        const nb = Number(b.match(/(\d+)\.xml$/)[1])
        return na - nb
      })
    for (const partPath of names) {
      const xml = zip.getText(partPath)
      layouts.push({
        index: layouts.length,
        partPath,
        name: layoutName(xml),
        type: getAttribute(rootOpeningTag(xml, 'p:sldLayout'), 'type') || 'custom',
        placeholders: layoutPlaceholders(xml)
      })
    }
  }

  return layouts
}

/** The human-facing layout name, `p:cSld/@name` falling back to the type. */
export function layoutName(layoutXml) {
  const cSld = openingTag(firstElement(layoutXml, 'p:cSld') || '')
  return getAttribute(cSld, 'name') || getAttribute(rootOpeningTag(layoutXml, 'p:sldLayout'), 'type') || 'Макет'
}

/**
 * Placeholder shapes declared by a layout or master, in document order.
 * @param {string} xml
 * @param {Map<string, object>|null} [extra] - the master's placeholder map, used
 *   when this part declares no transform of its own. R7's own title layout
 *   works that way: the position lives only in the master.
 * @returns {Array<object>}
 */
export function layoutPlaceholders(xml, extra = null) {
  const out = []
  for (const sp of extractElements(xml, 'p:sp')) {
    const ph = firstElement(sp.outerXml, 'p:ph')
    if (!ph) continue
    const nv = firstElement(sp.outerXml, 'p:cNvPr') || ''
    const xfrm = firstElement(sp.outerXml, 'a:xfrm')
    const off = xfrm ? firstElement(xfrm, 'a:off') : null
    const ext = xfrm ? firstElement(xfrm, 'a:ext') : null
    const type = getAttribute(ph, 'type') || 'body'
    const idx = getAttribute(ph, 'idx')
    let x = off ? Number(getAttribute(off, 'x') || 0) : null
    let y = off ? Number(getAttribute(off, 'y') || 0) : null
    let width = ext ? Number(getAttribute(ext, 'cx') || 0) : null
    let height = ext ? Number(getAttribute(ext, 'cy') || 0) : null

    if (width === null && extra) {
      const fallback = extra.get(`${type}|${idx === null ? '' : idx}`)
        || [...extra.values()].find((d) => d.type === type && d.width !== null)
      if (fallback) {
        x = x === null ? fallback.x : x
        y = y === null ? fallback.y : y
        width = fallback.width
        height = fallback.height
      }
    }

    out.push({
      type,
      idx,
      orient: getAttribute(ph, 'orient'),
      name: getAttribute(nv, 'name') || '',
      x, y, width, height
    })
  }
  return out
}

/** Presentation slide size in EMU. */
export function slideSize(zip) {
  const xml = zip.getText('ppt/presentation.xml')
  if (!xml) return { ...DEFAULT_SLIDE_SIZE }
  const sldSz = firstElement(xml, 'p:sldSz')
  if (!sldSz) return { ...DEFAULT_SLIDE_SIZE }
  const cx = Number(getAttribute(sldSz, 'cx') || 0)
  const cy = Number(getAttribute(sldSz, 'cy') || 0)
  if (!cx || !cy) return { ...DEFAULT_SLIDE_SIZE }
  return { width: cx, height: cy }
}

/**
 * Add an `<Override>` to `[Content_Types].xml` if the part is not registered.
 * @param {string} xml
 * @param {string} partName - absolute part name including the leading slash.
 * @param {string} contentType
 * @returns {string}
 */
export function ensureOverride(xml, partName, contentType) {
  if (xml.includes(`PartName="${partName}"`)) return xml
  return xml.replace(
    '</Types>',
    `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`
  )
}

/** The content type OOXML requires for each part kind the engine creates. */
export const CONTENT_TYPE = { ...CONTENT_TYPES }

/** Truncate a long text for a preview without splitting a surrogate pair. */
export function clip(text, max = 200) {
  const value = String(text ?? '')
  if (value.length <= max) return value
  return `${value.slice(0, max)}…`
}
