/**
 * sfnt.js — dependency-free TrueType/OpenType reader (Node stdlib only).
 *
 * Why this exists
 * ---------------
 * A CIDFontType2 PDF font is only correct when its glyph chain is consistent:
 *
 *     content code -> CID -> CIDToGIDMap[cid] -> GID -> glyf outline
 *     content code -> CID -> ToUnicode[cid]  -> Unicode -> font cmap -> GID
 *
 * Both chains must land on the same outline. They can be checked WITHOUT a
 * renderer, but only against the font the embedded subset was cut from: the
 * subset's own `cmap` is a private CID lookup (x2t writes a format 6 subtable
 * whose glyphIdArray is literally a copy of CIDToGIDMap, with a bogus `length`
 * field), so it can never name a glyph.
 *
 * This module reads just enough of an sfnt to (a) identify a font
 * (name/hmtx/maxp), (b) map Unicode to glyph ids (cmap 0/4/6/12) and
 * (c) fingerprint a glyph outline (glyf + loca, simple and composite) so two
 * glyphs can be compared by SHAPE rather than by index.
 */

const utf16be = (buf) => {
  let out = ''
  for (let i = 0; i + 1 < buf.length; i += 2) out += String.fromCharCode(buf.readUInt16BE(i))
  return out
}

/**
 * Read the sfnt table directory.
 * @param {Buffer} buf
 * @returns {{buf: Buffer, tables: Map<string, {offset: number, length: number}>}|null}
 */
export function readSfnt(buf) {
  if (!buf || buf.length < 12) return null
  if (buf.readUInt32BE(0) === 0x74746366) return null // 'ttcf' collections are not supported
  const numTables = buf.readUInt16BE(4)
  if (numTables < 1 || 12 + numTables * 16 > buf.length) return null
  const tables = new Map()
  for (let i = 0; i < numTables; i++) {
    const o = 12 + i * 16
    const tag = buf.toString('latin1', o, o + 4)
    const offset = buf.readUInt32BE(o + 8)
    const length = buf.readUInt32BE(o + 12)
    if (offset + length > buf.length) continue
    tables.set(tag, { offset, length })
  }
  return { buf, tables }
}

/**
 * Family/subfamily/PostScript name from the `name` table.
 * @returns {{family: string|null, subfamily: string|null, postScriptName: string|null}|null}
 */
export function readNameTable(sfnt) {
  const { buf, tables } = sfnt
  const name = tables.get('name')
  if (!name) return null
  const count = buf.readUInt16BE(name.offset + 2)
  const stringOffset = name.offset + buf.readUInt16BE(name.offset + 4)
  const out = { family: null, subfamily: null, postScriptName: null }
  const wanted = { 1: 'family', 2: 'subfamily', 16: 'family', 17: 'subfamily', 6: 'postScriptName' }
  const score = { family: 0, subfamily: 0, postScriptName: 0 }
  for (let i = 0; i < count; i++) {
    const o = name.offset + 6 + i * 12
    if (o + 12 > buf.length) break
    const platformID = buf.readUInt16BE(o)
    const nameID = buf.readUInt16BE(o + 6)
    const key = wanted[nameID]
    if (!key) continue
    const length = buf.readUInt16BE(o + 8)
    const offset = buf.readUInt16BE(o + 10)
    const start = stringOffset + offset
    if (start + length > buf.length) continue
    const raw = buf.subarray(start, start + length)
    // Platform 3 (Windows) is UTF-16BE; platform 1 (Mac) is a single-byte script.
    const value = platformID === 3 || platformID === 0 ? utf16be(raw) : raw.toString('latin1')
    if (!value) continue
    const rank = platformID === 3 ? 2 : 1
    // nameID 16/17 are "typographic" names and win over the legacy ones.
    if (rank + (nameID >= 16 ? 1 : 0) > score[key]) {
      score[key] = rank + (nameID >= 16 ? 1 : 0)
      out[key] = value
    }
  }
  return out
}

/** @returns {number} maxp.numGlyphs, or 0 when unavailable. */
export function readNumGlyphs(sfnt) {
  const maxp = sfnt.tables.get('maxp')
  if (!maxp || maxp.offset + 6 > sfnt.buf.length) return 0
  return sfnt.buf.readUInt16BE(maxp.offset + 4)
}

/** unitsPerEm from `head`, or 1000 when unavailable. */
export function readUnitsPerEm(sfnt) {
  const head = sfnt.tables.get('head')
  if (!head || head.offset + 20 > sfnt.buf.length) return 1000
  const upem = sfnt.buf.readUInt16BE(head.offset + 18)
  return upem > 0 ? upem : 1000
}

function readCmapSubtable(buf, off, format) {
  const map = new Map()
  if (format === 0) {
    for (let c = 0; c < 256; c++) {
      const g = buf.readUInt8(off + 6 + c)
      if (g) map.set(c, g)
    }
    return map
  }
  if (format === 4) {
    const segCountX2 = buf.readUInt16BE(off + 6)
    const segCount = segCountX2 / 2
    const endBase = off + 14
    const startBase = endBase + segCountX2 + 2
    const deltaBase = startBase + segCountX2
    const rangeBase = deltaBase + segCountX2
    if (rangeBase + segCountX2 > buf.length) return map
    for (let s = 0; s < segCount; s++) {
      const end = buf.readUInt16BE(endBase + s * 2)
      const start = buf.readUInt16BE(startBase + s * 2)
      const delta = buf.readInt16BE(deltaBase + s * 2)
      const rangeOffset = buf.readUInt16BE(rangeBase + s * 2)
      if (start === 0xffff) continue
      for (let c = start; c <= end && c <= 0xffff; c++) {
        let g
        if (rangeOffset === 0) g = (c + delta) & 0xffff
        else {
          const p = rangeBase + s * 2 + rangeOffset + (c - start) * 2
          if (p + 2 > buf.length) continue
          g = buf.readUInt16BE(p)
          if (g) g = (g + delta) & 0xffff
        }
        if (g) map.set(c, g)
      }
    }
    return map
  }
  if (format === 6) {
    const first = buf.readUInt16BE(off + 6)
    const count = buf.readUInt16BE(off + 8)
    for (let i = 0; i < count; i++) {
      const p = off + 10 + i * 2
      if (p + 2 > buf.length) break
      const g = buf.readUInt16BE(p)
      if (g) map.set(first + i, g)
    }
    return map
  }
  if (format === 12) {
    const groups = buf.readUInt32BE(off + 12)
    for (let i = 0; i < groups; i++) {
      const p = off + 16 + i * 12
      if (p + 12 > buf.length) break
      const start = buf.readUInt32BE(p)
      const end = buf.readUInt32BE(p + 4)
      const startGid = buf.readUInt32BE(p + 8)
      if (end < start || end - start > 0x10ffff) continue
      for (let c = start; c <= end; c++) map.set(c, startGid + (c - start))
    }
    return map
  }
  return map
}

/**
 * Unicode -> glyph id map, preferring the standard Unicode subtables.
 * (3,10) format 12 > (3,1) format 4 > (0,3) format 4 > (1,0) format 0/6.
 * @returns {Map<number, number>}
 */
export function readCmap(sfnt) {
  const { buf, tables } = sfnt
  const cmap = tables.get('cmap')
  if (!cmap) return new Map()
  const count = buf.readUInt16BE(cmap.offset + 2)
  const found = []
  for (let i = 0; i < count; i++) {
    const o = cmap.offset + 4 + i * 8
    if (o + 8 > buf.length) break
    const platformID = buf.readUInt16BE(o)
    const encodingID = buf.readUInt16BE(o + 2)
    const off = cmap.offset + buf.readUInt32BE(o + 4)
    if (off + 2 > buf.length) continue
    const format = buf.readUInt16BE(off)
    let rank = -1
    if (platformID === 3 && encodingID === 10 && format === 12) rank = 6
    else if (platformID === 3 && encodingID === 1 && format === 4) rank = 5
    else if (platformID === 0 && format === 4) rank = 4
    else if (platformID === 0 && format === 12) rank = 4
    else if (platformID === 1 && encodingID === 0 && (format === 0 || format === 6)) rank = 2
    if (rank < 0) continue
    found.push({ rank, off, format })
  }
  found.sort((a, b) => b.rank - a.rank)
  for (const sub of found) {
    const map = readCmapSubtable(buf, sub.off, sub.format)
    if (map.size) return map
  }
  return new Map()
}

/** hmtx advance in font units for one glyph. */
export function readAdvance(sfnt, gid) {
  const { buf, tables } = sfnt
  const hhea = tables.get('hhea')
  const hmtx = tables.get('hmtx')
  if (!hhea || !hmtx) return 0
  const numberOfHMetrics = buf.readUInt16BE(hhea.offset + 34)
  if (numberOfHMetrics === 0) return 0
  const index = gid < numberOfHMetrics ? gid : numberOfHMetrics - 1
  const p = hmtx.offset + index * 4
  if (p + 2 > buf.length) return 0
  return buf.readUInt16BE(p)
}

/** Raw bytes of a table, or null. */
export function tableBytes(sfnt, tag) {
  const t = sfnt.tables.get(tag)
  if (!t) return null
  return sfnt.buf.subarray(t.offset, t.offset + t.length)
}

/** Glyph offsets from loca (null entries mark empty glyphs). */
function readLoca(sfnt) {
  const { buf, tables } = sfnt
  const head = tables.get('head')
  const loca = tables.get('loca')
  const numGlyphs = readNumGlyphs(sfnt)
  if (!head || !loca || !numGlyphs) return null
  const longFormat = buf.readInt16BE(head.offset + 50) !== 0
  const offsets = new Array(numGlyphs + 1)
  for (let i = 0; i <= numGlyphs; i++) {
    const p = loca.offset + (longFormat ? i * 4 : i * 2)
    if (p + (longFormat ? 4 : 2) > buf.length) { offsets[i] = null; continue }
    offsets[i] = longFormat ? buf.readUInt32BE(p) : buf.readUInt16BE(p) * 2
  }
  return offsets
}

/** Glyph outline bytes for one glyph id, or null when the glyph is empty. */
export function readGlyph(sfnt, gid) {
  const glyf = sfnt.tables.get('glyf')
  const offsets = readLoca(sfnt)
  if (!glyf || !offsets) return null
  const start = offsets[gid]
  const end = offsets[gid + 1]
  if (start === null || end === null || end <= start) return null
  if (glyf.offset + end > sfnt.buf.length) return null
  return sfnt.buf.subarray(glyf.offset + start, glyf.offset + end)
}

/**
 * Decode a glyph outline into a canonical, encoding-independent description so
 * two fonts (or two re-encodings of the same outline) compare equal.
 * @returns {string|null}
 */
export function glyphOutlineKey(data, depth = 0) {
  if (!data || data.length < 10 || depth > 4) return null
  const numberOfContours = data.readInt16BE(0)
  if (numberOfContours >= 0) {
    let p = 10
    const ends = []
    for (let i = 0; i < numberOfContours; i++) {
      if (p + 2 > data.length) return null
      ends.push(data.readUInt16BE(p)); p += 2
    }
    const pointCount = numberOfContours === 0 ? 0 : ends[numberOfContours - 1] + 1
    if (p + 2 > data.length) return null
    const instructionLength = data.readUInt16BE(p)
    p += 2 + instructionLength
    if (p > data.length) return null
    const flags = []
    while (flags.length < pointCount) {
      if (p >= data.length) return null
      const f = data.readUInt8(p++); flags.push(f)
      if (f & 0x08) {
        if (p >= data.length) return null
        const repeat = data.readUInt8(p++)
        for (let k = 0; k < repeat && flags.length < pointCount; k++) flags.push(f)
      }
    }
    const xs = new Array(pointCount)
    let x = 0
    for (let i = 0; i < pointCount; i++) {
      const f = flags[i]
      if (f & 0x02) { if (p >= data.length) return null; const d = data.readUInt8(p++); x += (f & 0x10) ? d : -d }
      else if (!(f & 0x10)) { if (p + 2 > data.length) return null; x += data.readInt16BE(p); p += 2 }
      xs[i] = x
    }
    const ys = new Array(pointCount)
    let y = 0
    for (let i = 0; i < pointCount; i++) {
      const f = flags[i]
      if (f & 0x04) { if (p >= data.length) return null; const d = data.readUInt8(p++); y += (f & 0x20) ? d : -d }
      else if (!(f & 0x20)) { if (p + 2 > data.length) return null; y += data.readInt16BE(p); p += 2 }
      ys[i] = y
    }
    const points = []
    for (let i = 0; i < pointCount; i++) points.push(`${xs[i]},${ys[i]}${flags[i] & 1 ? 'o' : 'q'}`)
    return `s:${numberOfContours}:${ends.join('.')}:${points.join(' ')}`
  }
  // Composite: the component list plus each transform is the shape's identity.
  let p = 10
  const parts = []
  for (let guard = 0; guard < 64; guard++) {
    if (p + 4 > data.length) return null
    const flags = data.readUInt16BE(p); p += 2
    const glyphIndex = data.readUInt16BE(p); p += 2
    let a1, a2
    if (flags & 0x0001) {
      if (p + 4 > data.length) return null
      a1 = data.readInt16BE(p); a2 = data.readInt16BE(p + 2); p += 4
    } else {
      if (p + 2 > data.length) return null
      a1 = data.readInt8(p); a2 = data.readInt8(p + 1); p += 2
    }
    let transform = ''
    if (flags & 0x0008) { if (p + 2 > data.length) return null; transform = String(data.readInt16BE(p)); p += 2 }
    else if (flags & 0x0040) { if (p + 4 > data.length) return null; transform = `${data.readInt16BE(p)},${data.readInt16BE(p + 2)}`; p += 4 }
    else if (flags & 0x0080) {
      if (p + 8 > data.length) return null
      transform = `${data.readInt16BE(p)},${data.readInt16BE(p + 2)},${data.readInt16BE(p + 4)},${data.readInt16BE(p + 6)}`
      p += 8
    }
    parts.push(`${glyphIndex}[${a1},${a2}]${transform}`)
    if (!(flags & 0x0020)) break
  }
  return `c:${parts.join(' ')}`
}

/**
 * Best-effort identity of a font file: the fields that decide whether glyph ids
 * taken from one font mean the same thing in another.
 */
export function fontIdentity(buf) {
  const sfnt = readSfnt(buf)
  if (!sfnt) return null
  const name = readNameTable(sfnt)
  const numGlyphs = readNumGlyphs(sfnt)
  // Decoded per-glyph advances, not the raw `hmtx` bytes: two builds of the same
  // font routinely pack the same advances with a different long-metric count.
  const advances = Buffer.alloc(numGlyphs * 2)
  for (let gid = 0; gid < numGlyphs; gid++) advances.writeUInt16BE(readAdvance(sfnt, gid), gid * 2)
  return {
    sfnt,
    family: name ? name.family : null,
    subfamily: name ? name.subfamily : null,
    postScriptName: name ? name.postScriptName : null,
    numGlyphs,
    unitsPerEm: readUnitsPerEm(sfnt),
    advances,
    nameBytes: tableBytes(sfnt, 'name')
  }
}

/**
 * True when two font files are the same font build, so a glyph id taken from one
 * is the same glyph in the other. Every glyph's advance width must agree, which
 * pins the glyph order; the family name corroborates it.
 */
export function sameGlyphSpace(a, b) {
  if (!a || !b) return false
  if (a.numGlyphs !== b.numGlyphs || a.numGlyphs === 0) return false
  if (!a.advances || !b.advances || a.advances.length !== b.advances.length) return false
  if (Buffer.compare(a.advances, b.advances) !== 0) return false
  if (a.family && b.family && a.family !== b.family) return false
  return true
}
