/**
 * OOXML package plumbing shared by the DOCX feature modules.
 *
 * Everything the DOCX engine adds to a package — a media part, a header, a
 * hyperlink target — needs the same three things: a relationship in the
 * owning part's `.rels`, a content type in `[Content_Types].xml`, and a part
 * name that does not collide. This module owns those mechanics so the feature
 * modules never hand-roll package surgery.
 *
 * The edits are deliberately minimal string insertions rather than a parse
 * and re-serialization of the whole part: re-serializing `[Content_Types].xml`
 * or `word/_rels/document.xml.rels` would rewrite bytes the caller never asked
 * to change, and the project's ZIP layer goes out of its way to avoid exactly
 * that.
 */

import path from 'node:path'
import { extractElements, getAttribute } from '../shared/xml.js'

/** Base namespace of every officeDocument relationship type. */
export const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** Relationship types used by the DOCX features. */
export const REL_TYPES = {
  image: `${REL_BASE}/image`,
  hyperlink: `${REL_BASE}/hyperlink`,
  header: `${REL_BASE}/header`,
  footer: `${REL_BASE}/footer`,
  numbering: `${REL_BASE}/numbering`,
  styles: `${REL_BASE}/styles`
}

/** OOXML content types used by the DOCX features. */
export const CONTENT_TYPES = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  numbering: 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
}

/** Default content type per media extension. */
export const MEDIA_CONTENT_TYPES = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf'
}

/** Namespaces referenced by the drawing markup the engine writes. */
export const NS = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: REL_BASE
}

/** Package-absolute name of the main document part. */
export const DOCUMENT_PART = 'word/document.xml'

/** Package-absolute name of the main document relationships part. */
export const DOCUMENT_RELS_PART = 'word/_rels/document.xml.rels'

/**
 * Read one attribute from a fragment of an opening tag.
 * @param {string} xml
 * @param {string} name
 * @returns {string|null}
 */
export function attr(xml, name) {
  return getAttribute(xml, name)
}

/**
 * Parse a `.rels` part into a lookup structure.
 *
 * Relationship elements are self-closing in every producer seen in practice,
 * but a non-self-closing form is accepted so a hand-edited package still
 * parses.
 *
 * @param {string|null} xml
 * @returns {{list: Array<{id: string, type: string, target: string, targetMode: string|null}>, byId: Map<string, object>}}
 */
export function parseRelationships(xml) {
  const list = []
  if (!xml) return { list, byId: new Map() }

  const openRegex = /<Relationship\b([^>]*?)(\/?)>/gi
  let match
  while ((match = openRegex.exec(xml)) !== null) {
    const attrs = match[1]
    const id = attr(attrs, 'Id')
    if (!id) continue
    list.push({
      id,
      type: attr(attrs, 'Type') || '',
      target: attr(attrs, 'Target') || '',
      targetMode: attr(attrs, 'TargetMode') || null
    })
  }

  return { list, byId: new Map(list.map(r => [r.id, r])) }
}

/**
 * Allocate a relationship id that is not already used in the part.
 * @param {Array<{id: string}>} relationships
 * @returns {string}
 */
export function nextRelationshipId(relationships) {
  const used = new Set(relationships.map(r => r.id))
  let max = 0
  for (const rel of relationships) {
    const m = /^rId(\d+)$/.exec(rel.id)
    if (m) max = Math.max(max, Number(m[1]))
  }
  let candidate = `rId${max + 1}`
  while (used.has(candidate)) {
    max++
    candidate = `rId${max + 1}`
  }
  return candidate
}

/**
 * Append a relationship to a `.rels` part, returning the new XML and the id.
 *
 * @param {string} relsXml
 * @param {object} spec
 * @param {string} spec.type - relationship type URI
 * @param {string} spec.target - relationship target (relative or absolute)
 * @param {string} [spec.id] - explicit id; allocated when omitted
 * @param {string} [spec.targetMode] - `External` for hyperlinks
 * @returns {{xml: string, id: string, relationships: Array<object>}}
 */
export function addRelationship(relsXml, spec) {
  if (typeof relsXml !== 'string' || relsXml.length === 0) {
    throw new Error('Invalid DOCX: relationship part is missing or empty')
  }
  if (!spec || !spec.type || !spec.target) {
    throw new Error('addRelationship requires { type, target }')
  }

  const { list } = parseRelationships(relsXml)
  const id = spec.id || nextRelationshipId(list)
  if (list.some(r => r.id === id)) {
    throw new Error(`Relationship id "${id}" already exists in the part`)
  }

  const targetMode = spec.targetMode ? ` TargetMode="${spec.targetMode}"` : ''
  const entry = `<Relationship Id="${id}" Type="${spec.type}" Target="${escapeAttr(spec.target)}"${targetMode}/>`

  let xml
  if (relsXml.includes('</Relationships>')) {
    xml = relsXml.replace('</Relationships>', `${entry}</Relationships>`)
  } else {
    // Degenerate but legal form: an empty, self-closing root element.
    const selfClosing = /<Relationships\b([^>]*?)\/>/i.exec(relsXml)
    if (!selfClosing) {
      throw new Error('Invalid DOCX: relationship part has no <Relationships> root')
    }
    xml = relsXml.replace(selfClosing[0], `<Relationships${selfClosing[1]}>${entry}</Relationships>`)
  }

  return { xml, id, relationships: parseRelationships(xml).list }
}

/**
 * Remove a relationship by id. Returns the new XML and whether it existed.
 * @param {string} relsXml
 * @param {string} id
 * @returns {{xml: string, removed: boolean}}
 */
export function removeRelationship(relsXml, id) {
  const { list } = parseRelationships(relsXml)
  const entry = list.find(r => r.id === id)
  if (!entry) return { xml: relsXml, removed: false }

  // Remove exactly the element that parsed as this id: locate it by its own
  // serialized attribute signature rather than by a global id pattern, so a
  // neighbouring relationship with the id as a substring is never touched.
  const regex = new RegExp(
    `\\s*<Relationship\\b(?=[^>]*\\bId=["']${escapeRegExp(id)}["'])[^>]*?/>`,
    'i'
  )
  const xml = relsXml.replace(regex, '')
  return { xml, removed: xml !== relsXml }
}

/**
 * Parse `[Content_Types].xml`.
 * @param {string|null} xml
 * @returns {{defaults: Map<string, string>, overrides: Map<string, string>}}
 */
export function parseContentTypes(xml) {
  const defaults = new Map()
  const overrides = new Map()
  if (!xml) return { defaults, overrides }

  for (const match of xml.matchAll(/<Default\b([^>]*?)\/?>/gi)) {
    const ext = attr(match[1], 'Extension')
    const type = attr(match[1], 'ContentType')
    if (ext) defaults.set(ext.toLowerCase(), type || '')
  }
  for (const match of xml.matchAll(/<Override\b([^>]*?)\/?>/gi)) {
    const part = attr(match[1], 'PartName')
    const type = attr(match[1], 'ContentType')
    if (part) overrides.set(part, type || '')
  }
  return { defaults, overrides }
}

/**
 * Ensure a `<Default>` extension mapping exists, without rewriting one that
 * already does (a producer may map an extension to a different, equally valid
 * content type, and that decision is not ours to reverse).
 *
 * @param {string} xml
 * @param {string} extension - without the dot
 * @param {string} contentType
 * @returns {{xml: string, added: boolean, contentType: string}}
 */
export function ensureDefaultContentType(xml, extension, contentType) {
  const ext = String(extension).toLowerCase()
  const { defaults } = parseContentTypes(xml)
  if (defaults.has(ext)) {
    return { xml, added: false, contentType: defaults.get(ext) }
  }
  if (!xml.includes('</Types>')) {
    throw new Error('Invalid DOCX: [Content_Types].xml has no </Types> close tag')
  }
  const entry = `<Default Extension="${ext}" ContentType="${contentType}"/>`
  return {
    xml: xml.replace('</Types>', `${entry}</Types>`),
    added: true,
    contentType
  }
}

/**
 * Ensure an `<Override>` part-name mapping exists.
 * @param {string} xml
 * @param {string} partName - package-absolute, e.g. `/word/numbering.xml`
 * @param {string} contentType
 * @returns {{xml: string, added: boolean}}
 */
export function ensureOverrideContentType(xml, partName, contentType) {
  const name = partName.startsWith('/') ? partName : `/${partName}`
  const { overrides } = parseContentTypes(xml)
  if (overrides.has(name)) return { xml, added: false }
  if (!xml.includes('</Types>')) {
    throw new Error('Invalid DOCX: [Content_Types].xml has no </Types> close tag')
  }
  const entry = `<Override PartName="${name}" ContentType="${contentType}"/>`
  return { xml: xml.replace('</Types>', `${entry}</Types>`), added: true }
}

/**
 * Directory of a package part, POSIX-style.
 * @param {string} partName - e.g. `word/document.xml`
 * @returns {string}
 */
export function partDir(partName) {
  const idx = partName.lastIndexOf('/')
  return idx === -1 ? '' : partName.slice(0, idx)
}

/**
 * Resolve a relationship target against the part that owns it.
 * @param {string} ownerPart - e.g. `word/document.xml`
 * @param {string} target - e.g. `media/image1.png` or `/word/media/image1.png`
 * @returns {string} package-absolute part name without a leading slash
 */
export function resolveTarget(ownerPart, target) {
  if (!target) return ''
  if (target.startsWith('/')) return target.replace(/^\/+/, '')
  return path.posix.normalize(path.posix.join(partDir(ownerPart), target))
}

/**
 * Express a package part name relative to the part that references it.
 * @param {string} fromPart - e.g. `word/document.xml`
 * @param {string} toPart - e.g. `word/media/image1.png`
 * @returns {string}
 */
export function relativeTarget(fromPart, toPart) {
  const from = partDir(fromPart)
  const rel = path.posix.relative(from || '.', toPart)
  return rel === '' ? toPart : rel
}

/**
 * Find the next free part name for a numbered family, e.g. `word/media/image`.
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} prefix - package-relative prefix ending before the number
 * @param {string} extension - with or without the dot
 * @param {number} [start=1]
 * @returns {string}
 */
export function nextPartName(zip, prefix, extension, start = 1) {
  const ext = String(extension).replace(/^\./, '')
  const pattern = new RegExp(`^${escapeRegExp(prefix)}(\\d+)\\.${ext}$`, 'i')
  let max = start - 1
  for (const name of zip.list()) {
    const match = pattern.exec(name)
    if (match) max = Math.max(max, Number(match[1]))
  }
  return `${prefix}${max + 1}.${ext}`
}

/**
 * Package-absolute `/word/media/imageN.png` style name for a media part.
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} extension
 * @returns {string}
 */
export function nextMediaPartName(zip, extension) {
  return nextPartName(zip, 'word/media/image', extension)
}

/** Attribute-safe escaping (quotes included). */
export function escapeAttr(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/** Escape a string for use inside a regular expression. */
export function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Scan a well-formed XML fragment for its top-level child elements.
 *
 * OOXML property bags are shallow and well-formed, but rebuilding them with
 * regular expressions alone is how nested runs get mangled: `<w:r>` contains
 * `<w:rPr>` which contains `<w:rFonts>`, and a naive regex happily matches at
 * the wrong depth. This scanner tracks depth so callers get exact top-level
 * spans they can splice safely.
 *
 * @param {string} xml - fragment without its own root element (e.g. a `<w:p>` inner XML)
 * @returns {Array<{tag: string, start: number, end: number, xml: string, selfClosing: boolean}>}
 */
export function topLevelElements(xml) {
  const out = []
  if (typeof xml !== 'string' || xml.length === 0) return out

  const tagRegex = /<(\/?)([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g
  let depth = 0
  let currentStart = -1
  let currentTag = null
  let match

  while ((match = tagRegex.exec(xml)) !== null) {
    const closing = match[1] === '/'
    const selfClosing = match[4] === '/'
    if (closing) {
      depth--
      if (depth === 0 && currentStart !== -1) {
        out.push({
          tag: currentTag,
          start: currentStart,
          end: tagRegex.lastIndex,
          xml: xml.slice(currentStart, tagRegex.lastIndex),
          selfClosing: false
        })
        currentStart = -1
        currentTag = null
      }
      if (depth < 0) depth = 0
      continue
    }
    if (selfClosing) {
      if (depth === 0) {
        out.push({
          tag: match[2],
          start: match.index,
          end: tagRegex.lastIndex,
          xml: match[0],
          selfClosing: true
        })
      }
      continue
    }
    if (depth === 0) {
      currentStart = match.index
      currentTag = match[2]
    }
    depth++
  }

  return out
}

/**
 * Top-level children of a single-rooted element, with offsets relative to the
 * **whole** element.
 *
 * `topLevelElements` expects a fragment without its own root; passing a full
 * `<w:p>…</w:p>` to it yields one child (the paragraph itself) and silently
 * hides everything inside. This wrapper exists so callers holding a complete
 * element — which is what `extractElements` returns — can get the children
 * *and* splice them back using the same coordinates.
 *
 * @param {string} xml - one complete element, e.g. `<w:p>…</w:p>`
 * @returns {Array<{tag: string, start: number, end: number, xml: string, selfClosing: boolean}>}
 */
export function childElements(xml) {
  const openMatch = /^<([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/.exec(xml)
  if (!openMatch) return []
  if (openMatch[3] === '/') return []

  const tag = openMatch[1]
  const openTagLength = openMatch[0].length
  const closeTag = `</${tag}>`
  if (!xml.endsWith(closeTag)) return []

  const inner = xml.slice(openTagLength, xml.length - closeTag.length)
  return topLevelElements(inner).map(element => ({
    ...element,
    start: element.start + openTagLength,
    end: element.end + openTagLength
  }))
}

/**
 * Whether a character position lies inside the first element with a tag.
 * Useful to tell "the last paragraph" from "the last paragraph, inside a
 * table cell", which are very different insertion anchors.
 *
 * @param {string} xml
 * @param {string} tag
 * @param {number} position
 * @returns {boolean}
 */
export function isInsideElement(xml, tag, position) {
  for (const element of extractElements(xml, tag)) {
    if (position > element.index && position < element.index + element.outerXml.length) return true
  }
  return false
}

/**
 * Split one complete element into its open tag, inner XML and close tag.
 * @param {string} xml
 * @returns {{tag: string, open: string, inner: string, close: string, selfClosing: boolean}|null}
 */
export function elementParts(xml) {
  const match = /^<([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/.exec(xml)
  if (!match) return null
  const tag = match[1]
  if (match[3] === '/') {
    return { tag, open: match[0], inner: '', close: '', selfClosing: true }
  }
  const close = `</${tag}>`
  const end = xml.lastIndexOf(close)
  return {
    tag,
    open: match[0],
    inner: end === -1 ? xml.slice(match[0].length) : xml.slice(match[0].length, end),
    close: end === -1 ? '' : close,
    selfClosing: false
  }
}

/**
 * Rewrite a self-closing element into an open/close pair.
 *
 * `<w:p/>` is legal OOXML and shows up in header/footer parts and in
 * documents trimmed by other tools. Splicing children into it needs a real
 * close tag, and rebuilding it as `<w:p>…</w:p>` is what every producer does
 * anyway.
 *
 * @param {string} xml
 * @returns {string}
 */
export function expandSelfClosing(xml) {
  const parts = elementParts(xml)
  if (!parts) return xml
  if (!parts.selfClosing) return xml
  return `${parts.open.replace(/\/>$/, '>')}</${parts.tag}>`
}

/**
 * Remove every child element with the given tag from a container's inner XML.
 * @param {string} inner
 * @param {string} tag - qualified tag name, e.g. `w:tcW`
 * @param {{name: string, value: string}} [matchAttr] - only remove elements whose attribute matches
 * @returns {string}
 */
export function removeChildElements(inner, tag, matchAttr = null) {
  const matches = extractElements(inner, tag).filter(element => (
    !matchAttr || getAttribute(element.outerXml, matchAttr.name) === matchAttr.value
  ))
  let result = inner
  for (let i = matches.length - 1; i >= 0; i--) {
    const match = matches[i]
    result = result.slice(0, match.index) + result.slice(match.index + match.outerXml.length)
  }
  return result
}

/**
 * Insert a child element at its position in a schema order list.
 *
 * @param {string} inner - the container's inner XML
 * @param {string} tag - tag being inserted
 * @param {string} xml - the element
 * @param {string[]} order - every child tag, in schema order
 * @returns {string}
 */
export function insertChildOrdered(inner, tag, xml, order) {
  const tagIndex = order.indexOf(tag)
  if (tagIndex === -1) return inner + xml

  let insertAt = inner.length
  for (let i = tagIndex + 1; i < order.length; i++) {
    const element = extractElements(inner, order[i])[0]
    if (element) {
      insertAt = element.index
      break
    }
  }
  return inner.slice(0, insertAt) + xml + inner.slice(insertAt)
}

/**
 * Replace (or remove, when `xml` is null) a single-valued ordered child.
 * @param {string} inner
 * @param {string} tag
 * @param {string|null} xml
 * @param {string[]} order
 * @returns {string}
 */
export function upsertChildOrdered(inner, tag, xml, order) {
  const cleaned = removeChildElements(inner, tag)
  if (!xml) return cleaned
  return insertChildOrdered(cleaned, tag, xml, order)
}
