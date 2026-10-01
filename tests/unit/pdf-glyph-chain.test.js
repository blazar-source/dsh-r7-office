import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'

import { inspectPdf, inspectGlyphChain, analyzeGlyphChain, parseObjects, clearSourceFontCache } from '../../src/r7/pdf-inspect.js'
import { readSfnt, readCmap, readGlyph, glyphOutlineKey, fontIdentity, sameGlyphSpace } from '../../src/r7/sfnt.js'

// ---------------------------------------------------------------------------
// A tiny but structurally honest TrueType font, built here so the glyph-chain
// check can be exercised without a renderer, without R7 and without any system
// font. Two builds of the same description have the same glyph space, which is
// what the detector keys on.
// ---------------------------------------------------------------------------

function simpleGlyph(points) {
  const header = Buffer.alloc(10)
  header.writeInt16BE(1, 0) // one contour
  const xs = points.map((p) => p[0])
  const ys = points.map((p) => p[1])
  header.writeInt16BE(Math.min(...xs), 2)
  header.writeInt16BE(Math.min(...ys), 4)
  header.writeInt16BE(Math.max(...xs), 6)
  header.writeInt16BE(Math.max(...ys), 8)
  const ends = Buffer.alloc(2)
  ends.writeUInt16BE(points.length - 1, 0)
  const instructions = Buffer.alloc(2) // instructionLength = 0
  const flags = Buffer.from(points.map(() => 0x01)) // on-curve, explicit int16 deltas
  const xCoords = Buffer.alloc(points.length * 2)
  const yCoords = Buffer.alloc(points.length * 2)
  let px = 0
  let py = 0
  points.forEach((p, i) => {
    xCoords.writeInt16BE(p[0] - px, i * 2)
    yCoords.writeInt16BE(p[1] - py, i * 2)
    px = p[0]
    py = p[1]
  })
  return Buffer.concat([header, ends, instructions, flags, xCoords, yCoords])
}

const TRIANGLE = simpleGlyph([[0, 0], [100, 0], [50, 100]])
const SQUARE = simpleGlyph([[0, 0], [100, 0], [100, 100], [0, 100]])

function cmapFormat4(pairs) {
  // One contiguous segment per run, plus the mandatory 0xFFFF terminator.
  const segments = []
  for (const [start, gid] of pairs) segments.push({ start, end: start, delta: (gid - start) & 0xffff })
  segments.push({ start: 0xffff, end: 0xffff, delta: 1 })
  const segCount = segments.length
  const length = 14 + segCount * 8 + 2
  const buf = Buffer.alloc(length)
  buf.writeUInt16BE(4, 0)
  buf.writeUInt16BE(length, 2)
  buf.writeUInt16BE(0, 4)
  buf.writeUInt16BE(segCount * 2, 6)
  buf.writeUInt16BE(0, 8)
  buf.writeUInt16BE(0, 10)
  buf.writeUInt16BE(0, 12)
  let p = 14
  for (const s of segments) { buf.writeUInt16BE(s.end, p); p += 2 }
  buf.writeUInt16BE(0, p); p += 2
  for (const s of segments) { buf.writeUInt16BE(s.start, p); p += 2 }
  for (const s of segments) { buf.writeUInt16BE(s.delta & 0xffff, p); p += 2 }
  for (let i = 0; i < segCount; i++) { buf.writeUInt16BE(0, p); p += 2 }
  return buf
}

function nameTable(family, subfamily) {
  const strings = []
  const records = []
  const push = (nameID, text) => {
    const encoded = Buffer.from(text, 'utf16le').swap16()
    const offset = strings.reduce((sum, s) => sum + s.length, 0)
    strings.push(encoded)
    const record = Buffer.alloc(12)
    record.writeUInt16BE(3, 0)
    record.writeUInt16BE(1, 2)
    record.writeUInt16BE(0x409, 4)
    record.writeUInt16BE(nameID, 6)
    record.writeUInt16BE(encoded.length, 8)
    record.writeUInt16BE(offset, 10)
    records.push(record)
  }
  push(1, family)
  push(2, subfamily)
  const header = Buffer.alloc(6)
  header.writeUInt16BE(0, 0)
  header.writeUInt16BE(records.length, 2)
  header.writeUInt16BE(6 + records.length * 12, 4)
  return Buffer.concat([header, ...records, ...strings])
}

/**
 * Build a minimal sfnt.
 * @param {{family?: string, subfamily?: string, glyphs: Array<Buffer|null>, cmapPairs: Array<[number, number]>, advances: number[]}} spec
 */
function buildTtf({ family = 'Testfont', subfamily = 'Regular', glyphs, cmapPairs, advances }) {
  const numGlyphs = glyphs.length
  const glyfParts = []
  const offsets = [0]
  let offset = 0
  for (const glyph of glyphs) {
    const data = glyph || Buffer.alloc(0)
    const pad = (4 - (data.length % 4)) % 4
    const padded = Buffer.concat([data, Buffer.alloc(pad)])
    glyfParts.push(padded)
    offset += padded.length
    offsets.push(offset)
  }
  const glyf = Buffer.concat(glyfParts)
  const loca = Buffer.alloc((numGlyphs + 1) * 4)
  offsets.forEach((value, i) => loca.writeUInt32BE(value, i * 4))

  const head = Buffer.alloc(54)
  head.writeUInt32BE(0x00010000, 0)
  head.writeUInt32BE(0x5f0f3cf5, 12)
  head.writeUInt16BE(1000, 18) // unitsPerEm
  head.writeInt16BE(1, 50) // indexToLocFormat: long

  const hhea = Buffer.alloc(36)
  hhea.writeUInt32BE(0x00010000, 0)
  hhea.writeInt16BE(800, 4)
  hhea.writeInt16BE(-200, 6)
  hhea.writeUInt16BE(numGlyphs, 34) // numberOfHMetrics

  const maxp = Buffer.alloc(32)
  maxp.writeUInt32BE(0x00010000, 0)
  maxp.writeUInt16BE(numGlyphs, 4)

  const hmtx = Buffer.alloc(numGlyphs * 4)
  advances.forEach((advance, i) => hmtx.writeUInt16BE(advance, i * 4))

  const sub = cmapFormat4(cmapPairs)
  const cmap = Buffer.alloc(12)
  cmap.writeUInt16BE(0, 0)
  cmap.writeUInt16BE(1, 2)
  cmap.writeUInt16BE(3, 4)
  cmap.writeUInt16BE(1, 6)
  cmap.writeUInt32BE(12, 8)
  const cmapTable = Buffer.concat([cmap, sub])

  const tables = [
    ['cmap', cmapTable],
    ['glyf', glyf],
    ['head', head],
    ['hhea', hhea],
    ['hmtx', hmtx],
    ['loca', loca],
    ['maxp', maxp],
    ['name', nameTable(family, subfamily)]
  ]
  const numTables = tables.length
  const directory = Buffer.alloc(12 + numTables * 16)
  directory.writeUInt32BE(0x00010000, 0)
  directory.writeUInt16BE(numTables, 4)
  let dataOffset = 12 + numTables * 16
  const parts = [directory]
  tables.forEach(([tag, data], i) => {
    const entry = 12 + i * 16
    directory.write(tag, entry, 4, 'latin1')
    directory.writeUInt32BE(dataOffset, entry + 8)
    directory.writeUInt32BE(data.length, entry + 12)
    parts.push(data)
    const pad = (4 - (data.length % 4)) % 4
    if (pad) parts.push(Buffer.alloc(pad))
    dataOffset += data.length + pad
  })
  return Buffer.concat(parts)
}

// Two code points that share one outline (A -> GID 1, C -> GID 3 with GID 3 a
// copy of GID 1's triangle) plus B -> GID 2 (a square).
const GLYPHS = [null, TRIANGLE, SQUARE, TRIANGLE]
const CMAP_PAIRS = [[0x41, 1], [0x42, 2], [0x43, 3]]
const ADVANCES = [500, 600, 600, 600]

function writeTmpFile(name, contents) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-glyph-chain-'))
  const file = path.join(dir, name)
  fs.writeFileSync(file, contents)
  return file
}

function buildPdf(objects) {
  const parts = [Buffer.from('%PDF-1.7\n%\xE2\xE3\xCF\xD3\n', 'latin1')]
  for (const object of objects) {
    parts.push(Buffer.from(`${object.num} 0 obj\n`, 'latin1'))
    if (typeof object.body === 'string') {
      parts.push(Buffer.from(object.body + '\nendobj\n', 'latin1'))
    } else {
      parts.push(Buffer.from(object.body.dict + '\nstream\n', 'latin1'))
      parts.push(object.body.stream)
      parts.push(Buffer.from('\nendstream\nendobj\n', 'latin1'))
    }
  }
  parts.push(Buffer.from('trailer\n<< /Size 100 /Root 1 0 R >>\nstartxref\n0\n%%EOF\n', 'latin1'))
  return Buffer.concat(parts)
}

const deflate = (buf) => zlib.deflateSync(buf)

/**
 * A CIDFontType2 PDF whose CID->GID map is `cidToGid`, with a ToUnicode map that
 * names the real characters and an embedded subset cut from `font`.
 */
function cidPdf({ cidToGid, font, toUnicodePairs = [[1, '0041'], [2, '0042'], [3, '0043']] }) {
  const mapStream = Buffer.alloc(cidToGid.length * 2)
  cidToGid.forEach((gid, i) => mapStream.writeUInt16BE(gid, i * 2))
  const cmapText = [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '1 begincodespacerange',
    '<0000> <FFFF>',
    'endcodespacerange',
    `${toUnicodePairs.length} beginbfchar`,
    ...toUnicodePairs.map(([cid, hex]) => `<${cid.toString(16).padStart(4, '0')}> <${hex}>`),
    'endbfchar',
    'endcmap',
    'end',
    'end'
  ].join('\n')
  const content = 'BT\n/F1 12 Tf\n<000100020003> Tj\nET'
  const fontFile = deflate(font)
  const mapDeflated = deflate(mapStream)
  const toUnicodeDeflated = deflate(Buffer.from(cmapText, 'latin1'))
  const contentDeflated = deflate(Buffer.from(content, 'latin1'))
  return buildPdf([
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' },
    { num: 3, body: '<< /Type /Page /Parent 2 0 R /Contents [ 4 0 R ] /Resources << /Font << /F1 8 0 R >> >> >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${contentDeflated.length} >>`, stream: contentDeflated } },
    { num: 8, body: '<< /BaseFont /AAAAAA+Testfont /DescendantFonts [ 9 0 R ] /Encoding /Identity-H /Subtype /Type0 /ToUnicode 10 0 R /Type /Font >>' },
    { num: 9, body: '<< /BaseFont /AAAAAA+Testfont /CIDSystemInfo << /Ordering (Identity) /Registry (Adobe) /Supplement 0 >> /FontDescriptor 12 0 R /Subtype /CIDFontType2 /Type /Font /CIDToGIDMap 15 0 R >>' },
    { num: 10, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${toUnicodeDeflated.length} >>`, stream: toUnicodeDeflated } },
    { num: 12, body: '<< /Type /FontDescriptor /FontName /AAAAAA+Testfont /FontFile2 13 0 R >>' },
    { num: 13, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${fontFile.length} >>`, stream: fontFile } },
    { num: 15, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${mapDeflated.length} >>`, stream: mapDeflated } }
  ])
}

// ---------------------------------------------------------------------------
// sfnt reading
// ---------------------------------------------------------------------------

test('sfnt reads the table directory, cmap, glyph outlines and identity', () => {
  clearSourceFontCache()
  const ttf = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  const sfnt = readSfnt(ttf)
  assert.ok(sfnt, 'the built font must be a readable sfnt')
  for (const tag of ['cmap', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'name']) {
    assert.ok(sfnt.tables.has(tag), `missing table ${tag}`)
  }
  const cmap = readCmap(sfnt)
  assert.equal(cmap.get(0x41), 1)
  assert.equal(cmap.get(0x42), 2)
  assert.equal(cmap.get(0x43), 3)

  const identity = fontIdentity(ttf)
  assert.equal(identity.family, 'Testfont')
  assert.equal(identity.subfamily, 'Regular')
  assert.equal(identity.numGlyphs, 4)
  assert.equal(identity.unitsPerEm, 1000)
})

test('glyphOutlineKey distinguishes shapes and unifies identical outlines', () => {
  const ttf = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  const sfnt = readSfnt(ttf)
  const one = glyphOutlineKey(readGlyph(sfnt, 1))
  const two = glyphOutlineKey(readGlyph(sfnt, 2))
  const three = glyphOutlineKey(readGlyph(sfnt, 3))
  assert.ok(one && two && three)
  assert.notEqual(one, two, 'a triangle and a square must not compare equal')
  assert.equal(one, three, 'two glyphs holding the same outline must compare equal')
  assert.equal(glyphOutlineKey(readGlyph(sfnt, 0)), null, 'an empty glyph has no outline')
})

test('sameGlyphSpace accepts an identical font and rejects a different one', () => {
  const a = fontIdentity(buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES }))
  const b = fontIdentity(buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES }))
  const different = fontIdentity(buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: [500, 601, 600, 600] }))
  assert.equal(sameGlyphSpace(a, b), true)
  assert.equal(sameGlyphSpace(a, different), false, 'differing advances mean a different glyph space')
})

// ---------------------------------------------------------------------------
// The detector: "extracts fine, renders the wrong glyphs"
// ---------------------------------------------------------------------------

test('inspectGlyphChain accepts a PDF whose CID map matches the source font', () => {
  clearSourceFontCache()
  const font = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  const pdf = writeTmpFile('testfont-correct.pdf', cidPdf({ cidToGid: [0, 1, 2, 3], font }))
  const source = writeTmpFile('testfont.ttf', font)

  const chain = inspectGlyphChain(pdf, { fontPaths: [source] })
  assert.equal(chain.checked, true)
  assert.equal(chain.consistent, true)
  assert.equal(chain.cidsChecked, 3)
  assert.equal(chain.mismatches, 0)
  assert.equal(chain.sourceFontsMatched, 1)
})

test('inspectGlyphChain catches CIDs remapped to a different outline', () => {
  clearSourceFontCache()
  const font = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  // 'A' is glyph 1 but the map draws glyph 2 (a square): the classic substitute
  // defect, where the text layer still extracts the right characters.
  const pdf = writeTmpFile('testfont-broken.pdf', cidPdf({ cidToGid: [0, 2, 1, 3], font }))
  const source = writeTmpFile('testfont.ttf', font)

  const chain = inspectGlyphChain(pdf, { fontPaths: [source] })
  assert.equal(chain.checked, true)
  assert.equal(chain.consistent, false)
  assert.equal(chain.mismatches, 2)
  const font0 = chain.fonts[0]
  assert.equal(font0.baseFont, 'AAAAAA+Testfont')
  assert.equal(font0.sourceFont, source)
  assert.deepEqual(font0.examples[0], { cid: 1, character: 'A', drawnGid: 2, expectedGid: 1 })

  // And the same verdict through the public inspectPdf entry point.
  const report = inspectPdf(pdf, { fontPaths: [source] })
  assert.equal(report.glyphMapInconsistent, true)
  assert.equal(report.glyphChainMismatches, 2)
  assert.equal(report.glyphChainChecked, true)
  assert.equal(report.verdict, 'text', 'the text layer is untouched, which is the whole trap')
  assert.equal(report.textGlyphs, 3)
})

test('a CID pointing at a different glyph id with the SAME outline is not a defect', () => {
  clearSourceFontCache()
  const font = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  // GID 3 is a byte-identical copy of GID 1, exactly like a font that shares one
  // glyph between two look-alike characters (Arial does this for Cyrillic and
  // Greek). Comparing indices alone would report a false positive here.
  const pdf = writeTmpFile('testfont-homoglyph.pdf', cidPdf({ cidToGid: [0, 3, 2, 1], font }))
  const source = writeTmpFile('testfont.ttf', font)

  const chain = inspectGlyphChain(pdf, { fontPaths: [source] })
  assert.equal(chain.mismatches, 0)
  assert.equal(chain.consistent, true)
})

test('inspectGlyphChain reports "not checkable" without candidate fonts', () => {
  clearSourceFontCache()
  const font = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  const pdf = writeTmpFile('testfont-nofonts.pdf', cidPdf({ cidToGid: [0, 2, 1, 3], font }))

  const chain = inspectGlyphChain(pdf, { fontPaths: [] })
  assert.equal(chain.checked, false)
  assert.equal(chain.consistent, null)
  assert.equal(chain.mismatches, 0, 'an unverifiable PDF must never be reported as corrupt')
  assert.equal(chain.fonts[0].reason, 'no-matching-source-font')

  const report = inspectPdf(pdf)
  assert.equal(report.glyphChainChecked, false)
  assert.equal(report.glyphMapInconsistent, false)
})

test('inspectGlyphChain refuses to trust a different font version', () => {
  clearSourceFontCache()
  const font = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  const impostor = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: [500, 601, 600, 600] })
  const pdf = writeTmpFile('testfont-version.pdf', cidPdf({ cidToGid: [0, 2, 1, 3], font }))
  const source = writeTmpFile('testfont.ttf', impostor)

  const chain = inspectGlyphChain(pdf, { fontPaths: [source] })
  assert.equal(chain.checked, false)
  assert.equal(chain.consistent, null, 'a font with different metrics is not proof of anything')
})

test('inspectPdf tolerates a non-font FontFile2 stream', () => {
  clearSourceFontCache()
  const pdf = writeTmpFile('testfont-garbage.pdf', cidPdf({ cidToGid: [0, 1, 2, 3], font: Buffer.from('not a font at all') }))
  const source = writeTmpFile('testfont.ttf', buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES }))
  const chain = inspectGlyphChain(pdf, { fontPaths: [source] })
  assert.equal(chain.checked, false)
  assert.equal(chain.fonts[0].reason, 'unsupported-font-file')
})

test('analyzeGlyphChain folds its verdict into the font list', () => {
  const font = buildTtf({ glyphs: GLYPHS, cmapPairs: CMAP_PAIRS, advances: ADVANCES })
  const source = writeTmpFile('testfont.ttf', font)
  const parsed = parseObjects(cidPdf({ cidToGid: [0, 1, 2, 3], font }))
  const chain = analyzeGlyphChain(parsed, { fontPaths: [source] })
  assert.equal(chain.fonts.length, 1)
  assert.equal(chain.fonts[0].object, 8)
})
