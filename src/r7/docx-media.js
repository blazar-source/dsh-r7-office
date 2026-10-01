/**
 * DOCX images: `word/media/*`, the document relationship, the content type and
 * the inline `<w:drawing>` run.
 *
 * Inserting a picture into an OOXML package is four coordinated changes, and
 * three of them are easy to forget because a document with a drawing but no
 * media part still opens — it just shows a red placeholder:
 *
 *   1. the binary part (`word/media/imageN.ext`);
 *   2. a relationship in `word/_rels/document.xml.rels`;
 *   3. a `<Default>` content type for the extension, if the package has none;
 *   4. the `<w:drawing>` run inside a paragraph.
 *
 * Image dimensions are read from the file's own header (PNG IHDR, JPEG SOF,
 * GIF screen descriptor) so a caller can ask for "10 cm wide" and get a height
 * derived from the real aspect ratio instead of a guessed one.
 */

import fs from 'node:fs'
import path from 'node:path'
import {
  DOCUMENT_PART,
  DOCUMENT_RELS_PART,
  MEDIA_CONTENT_TYPES,
  REL_TYPES,
  addRelationship,
  elementParts,
  ensureDefaultContentType,
  expandSelfClosing,
  isInsideElement,
  nextMediaPartName,
  relativeTarget
} from './docx-parts.js'
import { appendToBodyEnd } from './docx-sections.js'
import { extractElements, getAttribute } from '../shared/xml.js'

/** English Metric Units per inch — the unit DrawingML uses for sizes. */
export const EMU_PER_INCH = 914400

/** EMU per CSS/bitmap pixel at 96 dpi. */
export const EMU_PER_PIXEL = 9525

/** EMU per centimetre. */
export const EMU_PER_CM = 360000

/** Convert pixels at 96 dpi to EMU. */
export function pixelsToEmu(pixels) {
  return Math.round(Number(pixels) * EMU_PER_PIXEL)
}

/** Convert centimetres to EMU. */
export function cmToEmu(cm) {
  return Math.round(Number(cm) * EMU_PER_CM)
}

/** Convert EMU to centimetres, rounded to two decimals. */
export function emuToCm(emu) {
  return Math.round((Number(emu) / EMU_PER_CM) * 100) / 100
}

/** Convert points to EMU. */
export function pointsToEmu(points) {
  return Math.round((Number(points) / 72) * EMU_PER_INCH)
}

/**
 * Read an image's intrinsic size from its header.
 *
 * @param {Buffer} buffer
 * @returns {{format: 'png'|'jpeg'|'gif'|null, pixelWidth: number|null, pixelHeight: number|null, extension: string|null, contentType: string|null}}
 */
export function probeImage(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 8) {
    return { format: null, pixelWidth: null, pixelHeight: null, extension: null, contentType: null }
  }

  // PNG: 8-byte signature, then the IHDR chunk whose payload starts at 16.
  if (buffer.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
    if (buffer.length >= 24 && buffer.toString('ascii', 12, 16) === 'IHDR') {
      return {
        format: 'png',
        pixelWidth: buffer.readUInt32BE(16),
        pixelHeight: buffer.readUInt32BE(20),
        extension: 'png',
        contentType: MEDIA_CONTENT_TYPES.png
      }
    }
    return { format: 'png', pixelWidth: null, pixelHeight: null, extension: 'png', contentType: MEDIA_CONTENT_TYPES.png }
  }

  // GIF: logical screen descriptor holds little-endian dimensions at 6 and 8.
  if (buffer.toString('ascii', 0, 4) === 'GIF8') {
    return {
      format: 'gif',
      pixelWidth: buffer.readUInt16LE(6),
      pixelHeight: buffer.readUInt16LE(8),
      extension: 'gif',
      contentType: MEDIA_CONTENT_TYPES.gif
    }
  }

  // JPEG: walk the marker segments until a start-of-frame marker appears.
  if (buffer[0] === 0xff && buffer[1] === 0xd8) {
    let offset = 2
    while (offset + 9 < buffer.length) {
      if (buffer[offset] !== 0xff) {
        offset++
        continue
      }
      const marker = buffer[offset + 1]
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        offset += 2
        continue
      }
      const length = buffer.readUInt16BE(offset + 2)
      const isStartOfFrame = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
      if (isStartOfFrame) {
        return {
          format: 'jpeg',
          pixelHeight: buffer.readUInt16BE(offset + 5),
          pixelWidth: buffer.readUInt16BE(offset + 7),
          extension: 'jpeg',
          contentType: MEDIA_CONTENT_TYPES.jpeg
        }
      }
      offset += 2 + length
    }
    return { format: 'jpeg', pixelWidth: null, pixelHeight: null, extension: 'jpeg', contentType: MEDIA_CONTENT_TYPES.jpeg }
  }

  return { format: null, pixelWidth: null, pixelHeight: null, extension: null, contentType: null }
}

/**
 * Compute a display size in EMU, preserving the aspect ratio unless the caller
 * named both dimensions explicitly.
 *
 * @param {object} options
 * @param {number} sourcePixelWidth
 * @param {number} sourcePixelHeight
 * @param {number} [options.widthCm]
 * @param {number} [options.heightCm]
 * @param {number} [options.widthPx]
 * @param {number} [options.heightPx]
 * @param {number} [options.widthEmu]
 * @param {number} [options.heightEmu]
 * @returns {{cx: number, cy: number, aspectRatio: number, scaled: 'both'|'width'|'height'|'natural'}}
 */
export function resolveDisplaySize(options, sourcePixelWidth, sourcePixelHeight) {
  const naturalWidth = sourcePixelWidth || 96
  const naturalHeight = sourcePixelHeight || 96
  const ratio = naturalHeight / naturalWidth

  const toEmu = (value, unit) => {
    if (value === undefined || value === null) return null
    if (unit === 'cm') return cmToEmu(value)
    if (unit === 'px') return pixelsToEmu(value)
    if (unit === 'emu') return Math.round(Number(value))
    if (unit === 'pt') return pointsToEmu(value)
    throw new Error(`Unknown size unit: ${unit}`)
  }

  const widthCandidates = [
    toEmu(options.widthEmu, 'emu'),
    toEmu(options.widthCm, 'cm'),
    toEmu(options.widthPx, 'px'),
    toEmu(options.widthPt, 'pt')
  ]
  const heightCandidates = [
    toEmu(options.heightEmu, 'emu'),
    toEmu(options.heightCm, 'cm'),
    toEmu(options.heightPx, 'px'),
    toEmu(options.heightPt, 'pt')
  ]

  const width = widthCandidates.find(value => value !== null) ?? null
  const height = heightCandidates.find(value => value !== null) ?? null

  if (width !== null && height !== null) {
    return { cx: width, cy: height, aspectRatio: width / height, scaled: 'both' }
  }
  if (width !== null) {
    const cy = Math.round(width * ratio)
    return { cx: width, cy, aspectRatio: width / cy, scaled: 'width' }
  }
  if (height !== null) {
    const cx = Math.round(height / ratio)
    return { cx, cy: height, aspectRatio: cx / height, scaled: 'height' }
  }

  return {
    cx: pixelsToEmu(naturalWidth),
    cy: pixelsToEmu(naturalHeight),
    aspectRatio: naturalWidth / naturalHeight,
    scaled: 'natural'
  }
}

/**
 * Build the `<w:drawing>` element for an inline picture.
 *
 * The `wp`, `a`, `pic` and `r` namespaces are declared on the subtree itself,
 * so the markup is valid inside a `word/document.xml` whose root declares only
 * `xmlns:w` (which is exactly what the engine's fallback template does).
 *
 * @param {object} spec
 * @param {string} spec.relId
 * @param {number} spec.cx - width in EMU
 * @param {number} spec.cy - height in EMU
 * @param {number} [spec.docPrId]
 * @param {string} [spec.name]
 * @param {string} [spec.description]
 * @returns {string}
 */
export function buildDrawingXml(spec) {
  const id = spec.docPrId ?? 1
  const name = escapeXmlAttr(spec.name || `Picture ${id}`)
  const description = spec.description ? ` descr="${escapeXmlAttr(spec.description)}"` : ''
  const nsWp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
  const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const nsPic = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  const nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

  return '<w:drawing>'
    + `<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${nsWp}">`
    + `<wp:extent cx="${spec.cx}" cy="${spec.cy}"/>`
    + '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    + `<wp:docPr id="${id}" name="${name}"${description}/>`
    + '<wp:cNvGraphicFramePr>'
    + `<a:graphicFrameLocks xmlns:a="${nsA}" noChangeAspect="1"/>`
    + '</wp:cNvGraphicFramePr>'
    + `<a:graphic xmlns:a="${nsA}">`
    + '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<pic:pic xmlns:pic="${nsPic}">`
    + `<pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/>`
    + '<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
    + '<pic:blipFill>'
    + `<a:blip xmlns:r="${nsR}" r:embed="${spec.relId}"/>`
    + '<a:stretch><a:fillRect/></a:stretch>'
    + '</pic:blipFill>'
    + '<pic:spPr>'
    + `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${spec.cx}" cy="${spec.cy}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    + '</pic:spPr>'
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
}

function escapeXmlAttr(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/**
 * Highest `wp:docPr/@id` in the document, so a new drawing gets a free id.
 * @param {string} docXml
 * @returns {number}
 */
export function maxDocPrId(docXml) {
  let max = 0
  for (const match of docXml.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/gi)) {
    max = Math.max(max, Number(match[1]))
  }
  return max
}

/**
 * Append a drawing run to a paragraph element.
 * @param {string} pXml
 * @param {string} drawingXml
 * @returns {string}
 */
export function appendDrawingRun(pXml, drawingXml) {
  const paragraphXml = expandSelfClosing(pXml)
  const parts = elementParts(paragraphXml)
  if (!parts || !parts.close) throw new Error('Cannot insert an image: the target is not a paragraph element')
  const end = paragraphXml.lastIndexOf(parts.close)
  return `${paragraphXml.slice(0, end)}<w:r>${drawingXml}</w:r>${paragraphXml.slice(end)}`
}

/**
 * Insert an image into a DOCX archive.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} docXml
 * @param {object} options
 * @param {Buffer} options.buffer - the image bytes
 * @param {string} [options.fileName] - used for the part name and default alt text
 * @param {number|'end'|'new'|'start'} [options.paragraphIndex='end']
 * @returns {{docXml: string, mediaPartName: string, relId: string, cx: number, cy: number, widthCm: number, heightCm: number, pixelWidth: number|null, pixelHeight: number|null, format: string|null, aspectRatio: number, scaled: string, paragraphIndex: number}}
 */
export function insertImage(zip, docXml, options) {
  let buffer = options.buffer
  if (!Buffer.isBuffer(buffer) && options.filePath) {
    buffer = fs.readFileSync(options.filePath)
  }
  if (!Buffer.isBuffer(buffer)) throw new Error('insertImage requires { buffer } or { filePath }')

  const probed = probeImage(buffer)
  const extension = options.extension
    ? String(options.extension).replace(/^\./, '').toLowerCase()
    : (path.extname(options.filePath || '').replace(/^\./, '').toLowerCase() || probed.extension || 'png')
  const contentType = options.contentType || MEDIA_CONTENT_TYPES[extension] || probed.contentType || 'application/octet-stream'

  const mediaPartName = nextMediaPartName(zip, extension)
  const relsPart = zip.getText(DOCUMENT_RELS_PART)
  const withRelationship = addRelationship(relsPart, {
    type: REL_TYPES.image,
    target: relativeTarget(DOCUMENT_PART, mediaPartName)
  })

  const contentTypes = zip.getText('[Content_Types].xml')
  let contentTypesXml = contentTypes
  if (contentTypes) {
    contentTypesXml = ensureDefaultContentType(contentTypes, extension, contentType).xml
  }

  zip.setBuffer(mediaPartName, buffer)
  zip.setText(DOCUMENT_RELS_PART, withRelationship.xml)
  if (contentTypesXml && contentTypesXml !== contentTypes) zip.setText('[Content_Types].xml', contentTypesXml)

  const size = resolveDisplaySize(options, probed.pixelWidth, probed.pixelHeight)
  const drawing = buildDrawingXml({
    relId: withRelationship.id,
    cx: size.cx,
    cy: size.cy,
    docPrId: maxDocPrId(docXml) + 1,
    name: options.name || (options.fileName ? path.basename(options.fileName) : `Picture ${maxDocPrId(docXml) + 1}`),
    description: options.alt || options.description || ''
  })

  const paragraphs = extractElements(docXml, 'w:p')
  const requested = options.paragraphIndex ?? 'end'
  let xml = docXml
  let targetIndex

  const newBodyParagraph = () => {
    const appended = appendToBodyEnd(docXml, `<w:p><w:r>${drawing}</w:r></w:p>`)
    xml = appended
    targetIndex = extractElements(xml, 'w:p').length - 1
  }

  if (requested === 'new') {
    newBodyParagraph()
  } else if (requested === 'end') {
    // "The end of the document" is not the last paragraph when that paragraph
    // lives inside a table cell; a picture belongs in the body flow instead.
    const last = paragraphs[paragraphs.length - 1]
    if (!last || isInsideElement(docXml, 'w:tbl', last.index)) {
      newBodyParagraph()
    } else {
      targetIndex = paragraphs.length - 1
      const updated = appendDrawingRun(last.outerXml, drawing)
      xml = docXml.slice(0, last.index) + updated + docXml.slice(last.index + last.outerXml.length)
    }
  } else if (requested === 'start') {
    targetIndex = 0
    const first = paragraphs[0]
    if (!first) {
      newBodyParagraph()
    } else {
      const updated = appendDrawingRun(first.outerXml, drawing)
      xml = docXml.slice(0, first.index) + updated + docXml.slice(first.index + first.outerXml.length)
    }
  } else {
    targetIndex = Number(requested)
    if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= paragraphs.length) {
      throw new Error(`Paragraph index out of range for an image: ${requested} (total: ${paragraphs.length})`)
    }
    const target = paragraphs[targetIndex]
    const updated = appendDrawingRun(target.outerXml, drawing)
    xml = docXml.slice(0, target.index) + updated + docXml.slice(target.index + target.outerXml.length)
  }

  return {
    docXml: xml,
    mediaPartName,
    relId: withRelationship.id,
    cx: size.cx,
    cy: size.cy,
    widthCm: emuToCm(size.cx),
    heightCm: emuToCm(size.cy),
    pixelWidth: probed.pixelWidth,
    pixelHeight: probed.pixelHeight,
    format: probed.format,
    aspectRatio: size.aspectRatio,
    scaled: size.scaled,
    paragraphIndex: targetIndex
  }
}

/**
 * Every image in the document, with its relationship and display size.
 * @param {string} docXml
 * @param {Map<string, object>|Array<object>} relationships
 * @returns {Array<object>}
 */
export function listImages(docXml, relationships) {
  const byId = relationships instanceof Map
    ? relationships
    : new Map((relationships || []).map(rel => [rel.id, rel]))
  const result = []
  const paragraphs = extractElements(docXml, 'w:p')

  for (let i = 0; i < paragraphs.length; i++) {
    const pXml = paragraphs[i].outerXml
    const drawings = extractElements(pXml, 'w:drawing')
    for (const drawing of drawings) {
      const extent = extractElements(drawing.outerXml, 'wp:extent')[0]
      const blip = extractElements(drawing.outerXml, 'a:blip')[0]
      const docPr = extractElements(drawing.outerXml, 'wp:docPr')[0]
      const relId = blip ? getAttribute(blip.outerXml, 'r:embed') : null
      const rel = relId ? byId.get(relId) : null
      const cx = extent ? Number(getAttribute(extent.outerXml, 'cx')) : null
      const cy = extent ? Number(getAttribute(extent.outerXml, 'cy')) : null
      result.push({
        paragraphIndex: i,
        relId,
        mediaPartName: rel ? normalizeMediaTarget(rel.target) : null,
        widthEmu: cx,
        heightEmu: cy,
        widthCm: cx === null ? null : emuToCm(cx),
        heightCm: cy === null ? null : emuToCm(cy),
        aspectRatio: cx && cy ? cx / cy : null,
        name: docPr ? getAttribute(docPr.outerXml, 'name') : null,
        description: docPr ? getAttribute(docPr.outerXml, 'descr') : null
      })
    }
  }

  return result
}

function normalizeMediaTarget(target) {
  if (!target) return null
  if (target.startsWith('/')) return target.replace(/^\/+/, '')
  if (target.startsWith('word/')) return target
  return `word/${target.replace(/^\.\//, '')}`
}
