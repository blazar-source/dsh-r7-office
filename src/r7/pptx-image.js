import fs from 'node:fs'
import path from 'node:path'
import { extractElements, getAttribute } from '../shared/xml.js'
import {
  REL,
  EMU_PER_PIXEL,
  ensureOverride,
  findElement,
  firstElement,
  nextRelId,
  openingTag,
  relsPathFor,
  setChild,
  toEmu,
  EMU_PER_INCH
} from './pptx-util.js'

/**
 * Image support for the PPTX engine.
 *
 * Everything is done with the standard library: PNG and JPEG headers are
 * parsed by hand for their pixel size, and the bytes are stored in the
 * package's own `ppt/media` directory through the ZIP layer's binary writer.
 * No image library is involved and no bytes are ever re-encoded.
 */

/** Content type OOXML requires for each media extension. */
const MEDIA_CONTENT_TYPES = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  jpg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
  webp: 'image/webp',
  svg: 'image/svg+xml'
}

/** Extensions R7 and PowerPoint accept for a picture part. */
export const SUPPORTED_IMAGE_EXTENSIONS = Object.keys(MEDIA_CONTENT_TYPES)

/**
 * Identify an image format from its magic bytes.
 * @param {Buffer} buffer
 * @returns {string|null} `png`, `jpeg`, `gif`, `bmp`, `tiff`, `webp` or `svg`
 */
export function sniffFormat(buffer) {
  if (!buffer || buffer.length < 4) return null
  if (buffer[0] === 0x89 && buffer[1] === 0x50 && buffer[2] === 0x4e && buffer[3] === 0x47) return 'png'
  if (buffer[0] === 0xff && buffer[1] === 0xd8) return 'jpeg'
  if (buffer.subarray(0, 3).toString('latin1') === 'GIF') return 'gif'
  if (buffer.subarray(0, 2).toString('latin1') === 'BM') return 'bmp'
  if (buffer[0] === 0x49 && buffer[1] === 0x49 && buffer[2] === 0x2a) return 'tiff'
  if (buffer[0] === 0x4d && buffer[1] === 0x4d && buffer[2] === 0x00) return 'tiff'
  if (buffer.subarray(0, 4).toString('latin1') === 'RIFF'
    && buffer.subarray(8, 12).toString('latin1') === 'WEBP') return 'webp'
  const head = buffer.subarray(0, 512).toString('utf8').trimStart()
  if (head.startsWith('<?xml') || head.startsWith('<svg')) return 'svg'
  return null
}

/**
 * The pixel dimensions of an image, from its header.
 *
 * @param {Buffer} buffer
 * @param {string} format - as returned by {@link sniffFormat}.
 * @returns {{width: number, height: number}|null}
 */
export function imageSize(buffer, format) {
  switch (format) {
    case 'png': return pngSize(buffer)
    case 'jpeg': return jpegSize(buffer)
    case 'gif':
      if (buffer.length < 10) return null
      return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) }
    case 'bmp':
      if (buffer.length < 26) return null
      return { width: buffer.readInt32LE(18), height: Math.abs(buffer.readInt32LE(22)) }
    case 'webp': return webpSize(buffer)
    case 'tiff': return tiffSize(buffer)
    default: return null
  }
}

/** PNG dimensions live in the IHDR chunk: 8 byte signature, then 13 bytes. */
function pngSize(buffer) {
  if (buffer.length < 24) return null
  if (buffer.subarray(12, 16).toString('latin1') !== 'IHDR') return null
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) }
}

/**
 * JPEG dimensions live in the first Start-Of-Frame segment, which is found by
 * walking the segment chain: every segment is a marker, a two-byte length and
 * its payload.
 */
function jpegSize(buffer) {
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
    if (marker === 0xd9 || marker === 0xda) return null
    const length = buffer.readUInt16BE(offset + 2)
    if (length < 2) return null
    const isSof = (marker >= 0xc0 && marker <= 0xc3)
      || (marker >= 0xc5 && marker <= 0xc7)
      || (marker >= 0xc9 && marker <= 0xcb)
      || (marker >= 0xcd && marker <= 0xcf)
    if (isSof) {
      return { height: buffer.readUInt16BE(offset + 5), width: buffer.readUInt16BE(offset + 7) }
    }
    offset += 2 + length
  }
  return null
}

/** WebP dimensions, for the three container layouts the format allows. */
function webpSize(buffer) {
  if (buffer.length < 30) return null
  const fourCC = buffer.subarray(12, 16).toString('latin1')
  if (fourCC === 'VP8X') {
    const width = 1 + (buffer[24] | (buffer[25] << 8) | (buffer[26] << 16))
    const height = 1 + (buffer[27] | (buffer[28] << 8) | (buffer[29] << 16))
    return { width, height }
  }
  if (fourCC === 'VP8 ') {
    return {
      width: buffer.readUInt16LE(26) & 0x3fff,
      height: buffer.readUInt16LE(28) & 0x3fff
    }
  }
  if (fourCC === 'VP8L') {
    const bits = buffer.readUInt32LE(21)
    return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 }
  }
  return null
}

/** TIFF dimensions, honouring the little-endian byte-order mark. */
function tiffSize(buffer) {
  const little = buffer[0] === 0x49
  const read16 = (at) => (little ? buffer.readUInt16LE(at) : buffer.readUInt16BE(at))
  const read32 = (at) => (little ? buffer.readUInt32LE(at) : buffer.readUInt32BE(at))
  if (buffer.length < 8) return null
  const ifd = read32(4)
  if (ifd + 2 > buffer.length) return null
  const count = read16(ifd)
  let width = null
  let height = null
  for (let i = 0; i < count; i++) {
    const entry = ifd + 2 + i * 12
    if (entry + 12 > buffer.length) break
    const tag = read16(entry)
    const type = read16(entry + 2)
    const value = type === 3 ? read16(entry + 8) : read32(entry + 8)
    if (tag === 256) width = value
    if (tag === 257) height = value
  }
  if (!width || !height) return null
  return { width, height }
}

/**
 * Read an image file from disk and describe it.
 * @param {string} filePath
 * @returns {{buffer: Buffer, format: string, extension: string, width: number|null, height: number|null, bytes: number}}
 */
export function loadImage(filePath) {
  if (!fs.existsSync(filePath)) throw new Error(`Image not found: ${filePath}`)
  const buffer = fs.readFileSync(filePath)
  if (buffer.length === 0) throw new Error(`Image is empty: ${filePath}`)
  const format = sniffFormat(buffer)
  if (!format) {
    throw new Error(
      `Unsupported image format for "${path.basename(filePath)}". `
      + `Supported: ${SUPPORTED_IMAGE_EXTENSIONS.join(', ')}.`
    )
  }
  const size = imageSize(buffer, format)
  return {
    buffer,
    format,
    extension: format === 'jpeg' ? 'jpg' : format,
    width: size ? size.width : null,
    height: size ? size.height : null,
    bytes: buffer.length
  }
}

/**
 * Work out the EMU extent for an image, preserving its aspect ratio unless the
 * caller pinned both sides.
 *
 * @param {{width: number|null, height: number|null}} natural
 * @param {object} spec - x, y, width, height, lockAspectRatio, scale.
 * @param {{width: number, height: number}} [slide]
 * @returns {{x: number, y: number, width: number, height: number}}
 */
export function fitImage(natural, spec = {}, slide = null) {
  const x = spec.x === undefined ? toEmu(0, 'x') : toEmu(spec.x, 'x')
  const y = spec.y === undefined ? toEmu(0, 'y') : toEmu(spec.y, 'y')

  const requestedWidth = spec.width === undefined ? null : toEmu(spec.width, 'width')
  const requestedHeight = spec.height === undefined ? null : toEmu(spec.height, 'height')
  const ratioLocked = spec.lockAspectRatio !== false

  if (natural.width && natural.height) {
    const ratio = natural.width / natural.height
    if (requestedWidth !== null && requestedHeight !== null) {
      if (!ratioLocked) return { x, y, width: requestedWidth, height: requestedHeight }
      // With both sides pinned and the ratio locked, the smaller scale wins so
      // the picture fits the box the caller asked for.
      const scale = Math.min(requestedWidth / (natural.width * EMU_PER_PIXEL), requestedHeight / (natural.height * EMU_PER_PIXEL))
      return {
        x, y,
        width: Math.round(natural.width * EMU_PER_PIXEL * scale),
        height: Math.round(natural.height * EMU_PER_PIXEL * scale)
      }
    }
    if (requestedWidth !== null) {
      return { x, y, width: requestedWidth, height: Math.round(requestedWidth / ratio) }
    }
    if (requestedHeight !== null) {
      return { x, y, width: Math.round(requestedHeight * ratio), height: requestedHeight }
    }
    if (spec.scale !== undefined) {
      const scale = Number(spec.scale)
      if (!Number.isFinite(scale) || scale <= 0) throw new Error(`scale must be positive, got "${spec.scale}"`)
      return {
        x, y,
        width: Math.round(natural.width * EMU_PER_PIXEL * scale),
        height: Math.round(natural.height * EMU_PER_PIXEL * scale)
      }
    }
    // Neither side given: place it at its natural size, scaled down to fit the
    // slide when it is larger than the canvas.
    let width = natural.width * EMU_PER_PIXEL
    let height = natural.height * EMU_PER_PIXEL
    if (slide) {
      const maxWidth = Math.round(slide.width * 0.72)
      const maxHeight = Math.round(slide.height * 0.62)
      if (width > maxWidth || height > maxHeight) {
        const scale = Math.min(maxWidth / width, maxHeight / height)
        width = Math.round(width * scale)
        height = Math.round(height * scale)
      }
    }
    return { x, y, width, height }
  }

  // Unknown natural size (an SVG, or a header this parser does not read):
  // fall back to an explicit or default box rather than guessing a ratio.
  return {
    x, y,
    width: requestedWidth !== null ? requestedWidth : Math.round(4 * EMU_PER_INCH),
    height: requestedHeight !== null ? requestedHeight : Math.round(3 * EMU_PER_INCH)
  }
}

/**
 * Store an image in the package and register its content type.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {Buffer} buffer
 * @param {string} extension
 * @returns {string} the new part path, e.g. `ppt/media/image3.png`
 */
export function storeMedia(zip, buffer, extension) {
  const names = zip.list()
  const used = new Set(names)
  let index = 1
  let partPath = `ppt/media/image${index}.${extension}`
  while (used.has(partPath)) {
    index++
    partPath = `ppt/media/image${index}.${extension}`
  }
  zip.setBuffer(partPath, buffer)
  registerContentType(zip, extension)
  return partPath
}

/** Make sure `[Content_Types].xml` has a Default for this extension. */
export function registerContentType(zip, extension) {
  const contentType = MEDIA_CONTENT_TYPES[String(extension).toLowerCase()]
  if (!contentType) throw new Error(`Unsupported image extension ".${extension}"`)
  const xml = zip.getText('[Content_Types].xml')
  if (!xml) throw new Error('Invalid PPTX: [Content_Types].xml not found')
  if (new RegExp(`Extension="${extension}"`, 'i').test(xml)) return
  zip.setText('[Content_Types].xml', xml.replace(
    '</Types>',
    `<Default Extension="${extension}" ContentType="${contentType}"/></Types>`
  ))
}

/**
 * Add an image relationship to a slide's relationship part.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {string} slidePartPath
 * @param {string} mediaPartPath
 * @returns {string} the relationship id
 */
export function addImageRelationship(zip, slidePartPath, mediaPartPath) {
  const relsPath = relsPathFor(slidePartPath)
  let relsXml = zip.getText(relsPath)
  if (!relsXml) {
    relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  }
  const relId = nextRelId(relsXml)
  const target = relativeTarget(slidePartPath, mediaPartPath)
  relsXml = relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${relId}" Type="${REL.image}" Target="${target}"/></Relationships>`
  )
  zip.setText(relsPath, relsXml)
  return relId
}

/**
 * Replace the bytes an existing image relationship points at.
 *
 * The relationship and the picture element are untouched, so nothing that
 * references the image — the slide, a layout, another slide sharing the same
 * media part — can be broken by the swap. Shared media is copied first, so a
 * replacement never changes an image somebody else is still using.
 *
 * @param {import('../shared/zip.js').ZipArchive} zip
 * @param {object} image - the `image` block from `readSlide`.
 * @param {object} loaded - the result of {@link loadImage}.
 * @returns {{partPath: string, replacedInPlace: boolean}}
 */
export function replaceMedia(zip, image, loaded) {
  if (!image || !image.mediaPath) throw new Error('This object has no image to replace')
  const extension = loaded.extension
  const existingExtension = path.extname(image.mediaPath).replace('.', '').toLowerCase()
  const normalised = existingExtension === 'jpeg' ? 'jpg' : existingExtension
  const shared = countMediaReferences(zip, image.mediaPath) > 1

  if (normalised === extension && !shared) {
    zip.setBuffer(image.mediaPath, loaded.buffer)
    registerContentType(zip, extension)
    return { partPath: image.mediaPath, replacedInPlace: true }
  }

  // A different format, or media another part also shows: a fresh part keeps
  // the other consumers exactly as they were.
  const partPath = storeMedia(zip, loaded.buffer, extension)
  return { partPath, replacedInPlace: false }
}

/** Point an existing image relationship at a different media part. */
export function retargetImageRelationship(zip, slidePartPath, relId, mediaPartPath) {
  const relsPath = relsPathFor(slidePartPath)
  const relsXml = zip.getText(relsPath)
  if (!relsXml) throw new Error(`Missing relationships part: ${relsPath}`)
  const rel = extractElements(relsXml, 'Relationship').find((r) => getAttribute(r.outerXml, 'Id') === relId)
  if (!rel) throw new Error(`Relationship ${relId} not found in ${relsPath}`)
  const target = relativeTarget(slidePartPath, mediaPartPath)
  const updated = rel.outerXml.replace(/Target="[^"]*"/, `Target="${target}"`)
  zip.setText(relsPath, relsXml.replace(rel.outerXml, updated))
}

/** How many relationship parts in the package point at this media part. */
export function countMediaReferences(zip, mediaPartPath) {
  let count = 0
  for (const name of zip.list()) {
    if (!name.endsWith('.rels')) continue
    const xml = zip.getText(name)
    if (!xml) continue
    const sourcePart = sourcePartOf(name)
    for (const rel of extractElements(xml, 'Relationship')) {
      const target = getAttribute(rel.outerXml, 'Target') || ''
      if (resolveRelTarget(sourcePart, target) === mediaPartPath) count++
    }
  }
  return count
}

/** Compute a relationship target relative to the declaring part's directory. */
export function relativeTarget(sourcePart, targetPart) {
  if (sourcePart === targetPart) throw new Error('A part cannot relate to itself')
  const from = sourcePart.split('/').slice(0, -1)
  const to = targetPart.split('/')
  let common = 0
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++
  const up = from.length - common
  const rest = to.slice(common).join('/')
  return `${'../'.repeat(up)}${rest}`
}

function resolveRelTarget(sourcePart, target) {
  const clean = String(target).split('#')[0]
  if (clean.startsWith('/')) return clean.slice(1)
  const base = sourcePart.split('/').slice(0, -1)
  const stack = [...base]
  for (const segment of clean.split('/')) {
    if (!segment || segment === '.') continue
    if (segment === '..') stack.pop()
    else stack.push(segment)
  }
  return stack.join('/')
}

/** The part a `.rels` file describes. */
function sourcePartOf(relsName) {
  const dir = relsName.split('/').slice(0, -1)
  const file = relsName.split('/').pop().replace(/\.rels$/, '')
  const parent = dir[dir.length - 1] === '_rels' ? dir.slice(0, -1) : dir
  return `${parent.join('/')}${parent.length > 0 ? '/' : ''}${file}`
}

/**
 * Insert a picture into a slide's shape tree.
 *
 * @param {string} slideXml
 * @param {string} pictureXml
 * @returns {string}
 */
export function insertPicture(slideXml, pictureXml) {
  const tree = findElement(slideXml, 'p:spTree')
  if (!tree) throw new Error('Invalid slide: <p:spTree> not found')
  const inner = `${tree.innerXml}${pictureXml}`
  const updated = tree.outerXml.replace(tree.innerXml, inner)
  return slideXml.slice(0, tree.start) + updated + slideXml.slice(tree.end)
}

/**
 * Update a picture's geometry in place, keeping its relationship.
 * @param {string} pictureXml
 * @param {object} spec - x, y, width, height, rotation, name, description.
 * @returns {string}
 */
export function patchPicture(pictureXml, spec) {
  const spPr = findElement(pictureXml, 'p:spPr')
  if (!spPr) return pictureXml
  const xfrm = firstElement(spPr.outerXml, 'a:xfrm')
  const off = xfrm ? firstElement(xfrm, 'a:off') : null
  const ext = xfrm ? firstElement(xfrm, 'a:ext') : null
  const currentX = off ? Number(getAttribute(openingTag(off), 'x') || 0) : 0
  const currentY = off ? Number(getAttribute(openingTag(off), 'y') || 0) : 0
  const currentCx = ext ? Number(getAttribute(openingTag(ext), 'cx') || 0) : 0
  const currentCy = ext ? Number(getAttribute(openingTag(ext), 'cy') || 0) : 0

  const built = '<a:xfrm>'
    + `<a:off x="${spec.x === undefined ? currentX : toEmu(spec.x, 'x')}" y="${spec.y === undefined ? currentY : toEmu(spec.y, 'y')}"/>`
    + `<a:ext cx="${spec.width === undefined ? currentCx : toEmu(spec.width, 'width')}" cy="${spec.height === undefined ? currentCy : toEmu(spec.height, 'height')}"/>`
    + '</a:xfrm>'
  let pr = setChild(spPr.outerXml, 'a:xfrm', built, ['a:xfrm', 'a:custGeom', 'a:prstGeom', 'a:noFill', 'a:solidFill', 'a:ln', 'a:effectLst', 'a:extLst'])

  if (spec.name || spec.description) {
    const cNvPr = firstElement(pr, 'p:cNvPr')
    if (cNvPr) {
      let updated = cNvPr
      if (spec.name) updated = updated.replace(/name="[^"]*"/, `name="${spec.name.replace(/"/g, '&quot;')}"`)
      if (spec.description) {
        updated = /descr="/.test(updated)
          ? updated.replace(/descr="[^"]*"/, `descr="${spec.description.replace(/"/g, '&quot;')}"`)
          : updated.replace(/\/>$/, ` descr="${spec.description.replace(/"/g, '&quot;')}"/>`)
      }
      pr = pr.replace(cNvPr, updated)
    }
  }

  return pictureXml.replace(spPr.outerXml, pr)
}

/** Remove an image relationship whose only consumer is gone. */
export function dropImageRelationship(zip, slidePartPath, relId) {
  const relsPath = relsPathFor(slidePartPath)
  const relsXml = zip.getText(relsPath)
  if (!relsXml) return false
  const rel = extractElements(relsXml, 'Relationship').find((r) => getAttribute(r.outerXml, 'Id') === relId)
  if (!rel) return false
  zip.setText(relsPath, relsXml.replace(rel.outerXml, ''))
  return true
}

/** Every media part in the package, with its format and byte size. */
export function listMedia(zip) {
  const out = []
  for (const name of zip.list()) {
    if (!/^ppt\/media\//.test(name)) continue
    const buffer = zip.getBuffer(name)
    if (!buffer) continue
    const format = sniffFormat(buffer)
    const size = imageSize(buffer, format)
    out.push({
      partPath: name,
      format,
      bytes: buffer.length,
      width: size ? size.width : null,
      height: size ? size.height : null,
      referenced: countMediaReferences(zip, name) > 0
    })
  }
  return out
}

/** Register a media content type explicitly, used by the packages it creates. */
export function ensureMediaOverride(zip, partPath, extension) {
  const xml = zip.getText('[Content_Types].xml')
  const contentType = MEDIA_CONTENT_TYPES[extension]
  if (!xml || !contentType) return
  zip.setText('[Content_Types].xml', ensureOverride(xml, `/${partPath}`, contentType))
}
