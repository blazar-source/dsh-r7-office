import { test } from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import zlib from 'node:zlib'

import {
  inspectPdf,
  analyzeContentStream,
  parseToUnicodeCMap,
  findMalformedCMapBlocks,
  repairToUnicodeCMaps,
  parseObjects,
  isCyrillicCodePoint
} from '../../src/r7/pdf-inspect.js'

// ---------------------------------------------------------------------------
// Fixture helpers: build small but structurally honest PDFs, no dependencies.
// ---------------------------------------------------------------------------

let fileCounter = 0
function writeTmpPdf(buffer) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-inspect-test-'))
  const file = path.join(dir, `fixture-${fileCounter++}.pdf`)
  fs.writeFileSync(file, buffer)
  return file
}

/**
 * objects: array of { num, body } where body is either a dictionary string or
 * { dict, stream: Buffer }.
 */
function buildPdf(objects, { trailer = '<< /Size 100 /Root 1 0 R >>' } = {}) {
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
  parts.push(Buffer.from(`trailer\n${trailer}\nstartxref\n0\n%%EOF\n`, 'latin1'))
  return Buffer.concat(parts)
}

function deflate(text) {
  return zlib.deflateSync(Buffer.from(text, 'latin1'))
}

/** A CID-keyed Type0 font with an embedded descendant and a ToUnicode CMap. */
function fontObjects({ cmapPairs, declaredCount, fontFile = 'AAAA' }) {
  return [
    { num: 8, body: '<<\n/BaseFont /AAAAAA+Arial\n/DescendantFonts [ 9 0 R ]\n/Encoding /Identity-H\n/Subtype /Type0\n/ToUnicode 10 0 R\n/Type /Font\n>>' },
    { num: 9, body: '<<\n/BaseFont /AAAAAA+Arial\n/CIDSystemInfo << /Ordering (Identity) /Registry (Adobe) /Supplement 0 >>\n/FontDescriptor 12 0 R\n/Subtype /CIDFontType2\n/Type /Font\n>>' },
    {
      num: 10,
      body: {
        dict: '<<\n/Filter [ /FlateDecode ]\n/Length 11 0 R\n>>',
        stream: deflate(cmapText(cmapPairs, declaredCount))
      }
    },
    { num: 11, body: String(deflate(cmapText(cmapPairs, declaredCount)).length) },
    { num: 12, body: '<<\n/Type /FontDescriptor\n/FontName /AAAAAA+Arial\n/FontFile2 13 0 R\n>>' },
    { num: 13, body: { dict: '<<\n/Filter [ /FlateDecode ]\n/Length 14 0 R\n>>', stream: deflate('not-a-real-font ' + fontFile) } },
    { num: 14, body: String(deflate('not-a-real-font ' + fontFile).length) }
  ]
}

function cmapText(pairs, declaredCount) {
  const count = declaredCount === undefined ? pairs.length : declaredCount
  return [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<0000> <FFFF>',
    'endcodespacerange',
    `${count} beginbfchar`,
    ...pairs.map(([code, hex]) => `<${code}> <${hex}>`),
    'endbfchar',
    'endcmap',
    'end',
    'end'
  ].join('\n')
}

const CYRILLIC_PAIRS = [
  ['0001', '041E'], // О
  ['0004', '043D'], // н
  ['0007', '0430'], // а
  ['000C', '0440'], // р
  ['000E', '0438']  // и
]

// ---------------------------------------------------------------------------
// Content stream analysis
// ---------------------------------------------------------------------------

test('analyzeContentStream counts glyphs drawn by Tj and TJ', () => {
  const stream = 'q\nBT\n/F1 12 Tf\n10 10 Td\n<000100020003> Tj\nET\nQ\nBT\n/F1 12 Tf\n[(ab) -20 <00040005>] TJ\nET\nQ'
  const result = analyzeContentStream(stream)
  // 3 hex glyphs, then 2 literal + 2 hex glyphs.
  assert.equal(result.textGlyphs, 7)
  assert.equal(result.textRuns, 2)
})

test('analyzeContentStream reports fills that had no path to paint', () => {
  // Exactly the shape R7 emits when the renderer had no font to draw with.
  const broken = 'q\n0 0 0 rg\n/E1 gs\nf\nQ\n'.repeat(60)
  const brokenResult = analyzeContentStream(broken)
  assert.equal(brokenResult.textGlyphs, 0)
  assert.equal(brokenResult.emptyFillOperators, 60)

  // A real outlined glyph builds a path first: not an empty fill.
  const outlined = 'q\n100 100 m\n120 100 l\n120 140 l\n100 140 l\nh\nf\nQ'
  const outlinedResult = analyzeContentStream(outlined)
  assert.equal(outlinedResult.emptyFillOperators, 0)
  assert.ok(outlinedResult.pathOperators >= 4)
})

test('analyzeContentStream does not count text operators as path operators', () => {
  const stream = 'BT\n/F1 12 Tf\n10 10 Td\n<0001> Tj\nET'
  const result = analyzeContentStream(stream)
  assert.equal(result.pathOperators, 0)
  assert.equal(result.emptyFillOperators, 0)
  assert.equal(result.textGlyphs, 1)
})

test('analyzeContentStream skips inline images', () => {
  // The `f` inside the image payload must not be counted as a glyph fill.
  const stream = 'q\nBI /W 4 /H 4 ID \x00\x01 f Q m l EI\nQ'
  const result = analyzeContentStream(stream)
  assert.equal(result.emptyFillOperators, 0)
  assert.equal(result.textGlyphs, 0)
})

// ---------------------------------------------------------------------------
// ToUnicode CMap parsing
// ---------------------------------------------------------------------------

test('parseToUnicodeCMap reads bfchar pairs', () => {
  const pairs = parseToUnicodeCMap(cmapText(CYRILLIC_PAIRS))
  assert.equal(pairs.length, 5)
  assert.equal(pairs[0].code, 1)
  assert.equal(pairs[0].unicode, 'О')
  assert.ok(pairs.every((pair) => isCyrillicCodePoint(pair.unicode.codePointAt(0))))
})

test('parseToUnicodeCMap reads both bfrange forms', () => {
  const text = [
    '1 beginbfrange',
    '<0010> <0012> <0410>',
    'endbfrange',
    '1 beginbfrange',
    '<0020> <0021> [<0411> <0412>]',
    'endbfrange'
  ].join('\n')
  const pairs = parseToUnicodeCMap(text)
  assert.deepEqual(pairs.map((pair) => pair.unicode), ['А', 'Б', 'В', 'Б', 'В'])
})

test('findMalformedCMapBlocks detects a wrong declared count', () => {
  const problems = findMalformedCMapBlocks(cmapText(CYRILLIC_PAIRS, 944))
  assert.deepEqual(problems, [{ keyword: 'beginbfchar', declared: 944, actual: 5 }])
  assert.deepEqual(findMalformedCMapBlocks(cmapText(CYRILLIC_PAIRS)), [])
})

// ---------------------------------------------------------------------------
// inspectPdf
// ---------------------------------------------------------------------------

test('inspectPdf reports real text with Cyrillic as verdict "text"', () => {
  // A normal page: some drawn geometry (a table border) plus a text run.
  const content = [
    'q', '100 100 m', '200 100 l', '200 200 l', '100 200 l', 'h', 'f', 'Q',
    'BT', '/F1 12 Tf', '10 10 Td', '[<000100040007> -20 <000C000E>] TJ', 'ET'
  ].join('\n')
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' },
    { num: 3, body: '<< /Type /Page /Parent 2 0 R /Contents [ 4 0 R ] /Resources << /Font << /F1 8 0 R >> >> >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${deflate(content).length} >>`, stream: deflate(content) } },
    ...fontObjects({ cmapPairs: CYRILLIC_PAIRS })
  ]
  const file = writeTmpPdf(buildPdf(objects))
  const report = inspectPdf(file)

  assert.equal(report.pages, 1)
  assert.equal(report.textGlyphs, 5)
  assert.equal(report.emptyFillOperators, 0)
  assert.equal(report.hasExtractableText, true)
  assert.equal(report.hasCyrillicText, true)
  assert.equal(report.cyrillicMappings, 5)
  assert.equal(report.verdict, 'text')
  assert.equal(report.textLossSuspected, false)
  assert.equal(report.fonts.length, 1, 'descendant fonts must be folded into their parent')
  assert.equal(report.fonts[0].baseFont, 'AAAAAA+Arial')
  assert.equal(report.fonts[0].subtype, 'Type0')
  assert.equal(report.fonts[0].descendantSubtype, 'CIDFontType2')
  assert.equal(report.fonts[0].embedded, true)
  assert.equal(report.fonts[0].toUnicodeMappings, 5)
  assert.equal(report.fonts[0].cyrillicMappings, 5)
})

test('inspectPdf reports verdict "mixed" when text is drawn beside empty glyph fills', () => {
  const content = 'BT\n/F1 12 Tf\n<000100040007000C000E> Tj\nET\n' + 'q\n0 0 0 rg\n/E1 gs\nf\nQ\n'.repeat(200)
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' },
    { num: 3, body: '<< /Type /Page /Parent 2 0 R /Contents [ 4 0 R ] /Resources << /Font << /F1 8 0 R >> >> >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${deflate(content).length} >>`, stream: deflate(content) } },
    ...fontObjects({ cmapPairs: CYRILLIC_PAIRS })
  ]
  const report = inspectPdf(writeTmpPdf(buildPdf(objects)))
  assert.equal(report.textGlyphs, 5)
  assert.equal(report.emptyFillOperators, 200)
  assert.equal(report.verdict, 'mixed')
  assert.equal(report.textLossSuspected, true)
})

test('inspectPdf flags a PDF whose glyph fills painted nothing', () => {
  const content = 'q\n0 0 0 rg\n/E1 gs\nf\nQ\n'.repeat(400)
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 5 0 R >>' },
    { num: 3, body: '<< /Type /Page /Parent 5 0 R /Contents [ 4 0 R ] >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${deflate(content).length} >>`, stream: deflate(content) } },
    { num: 5, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' }
  ]
  const report = inspectPdf(writeTmpPdf(buildPdf(objects)))

  assert.equal(report.textGlyphs, 0)
  assert.equal(report.emptyFillOperators, 400)
  assert.equal(report.hasExtractableText, false)
  assert.equal(report.verdict, 'outlined')
  assert.equal(report.textLossSuspected, true)
})

test('inspectPdf does not flag an image-only page as text loss', () => {
  const content = 'q\n595 0 0 842 0 0 cm\n/Im1 Do\nQ'
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 5 0 R >>' },
    { num: 3, body: '<< /Type /Page /Parent 5 0 R /Contents [ 4 0 R ] >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${deflate(content).length} >>`, stream: deflate(content) } },
    { num: 5, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' }
  ]
  const report = inspectPdf(writeTmpPdf(buildPdf(objects)))
  assert.equal(report.emptyFillOperators, 0)
  assert.equal(report.textLossSuspected, false)
})

test('inspectPdf accepts a Buffer as well as a path', () => {
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 0 /Kids [] >>' }
  ]
  const report = inspectPdf(buildPdf(objects))
  assert.equal(report.filePath, null)
  assert.equal(report.pages, 0)
})

// ---------------------------------------------------------------------------
// /Length handling — the failure mode that hides most objects in a PDF
// ---------------------------------------------------------------------------

test('parseObjects honours a multi-digit indirect /Length written before the stream', () => {
  const content = `BT\n/F1 12 Tf\n<0001> Tj\nET\n` + 'q\nf\nQ\n'.repeat(30)
  const compressed = deflate(content)
  assert.ok(compressed.length >= 10, 'fixture length must be multi-digit for this regression')

  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 0 /Kids [] >>' },
    // Length object FIRST, so the indirect value is what must be used.
    { num: 6, body: String(compressed.length) },
    { num: 5, body: { dict: '<< /Filter [ /FlateDecode ] /Length 6 0 R >>', stream: compressed } }
  ]
  const parsed = parseObjects(buildPdf(objects))
  const stream = parsed.objects.get(5)
  assert.equal(stream.raw.length, compressed.length)
  const decoded = zlib.inflateSync(stream.raw).toString('latin1')
  assert.equal(decoded, content, 'a naive /Length regex would truncate this stream')
})

test('parseObjects falls back safely when the length object comes after the stream', () => {
  const content = `BT\n/F1 12 Tf\n<0001> Tj\nET\n` + 'q\nf\nQ\n'.repeat(30)
  const compressed = deflate(content)
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 0 /Kids [] >>' },
    { num: 5, body: { dict: '<< /Filter [ /FlateDecode ] /Length 6 0 R >>', stream: compressed } },
    { num: 6, body: String(compressed.length) }
  ]
  const parsed = parseObjects(buildPdf(objects))
  assert.equal(parsed.objects.get(5).raw.length, compressed.length)
  assert.equal(zlib.inflateSync(parsed.objects.get(5).raw).toString('latin1'), content)
})

test('parseObjects does not split a stream on a literal "endstream" inside it', () => {
  const content = 'BT\n/F1 12 Tf\n<0001> Tj\n(endstream is just text) Tj\nET\nf'
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 0 /Kids [] >>' },
    { num: 7, body: String(Buffer.byteLength(content, 'latin1')) },
    { num: 5, body: { dict: '<< /Length 7 0 R >>', stream: Buffer.from(content, 'latin1') } }
  ]
  const parsed = parseObjects(buildPdf(objects))
  assert.equal(parsed.objects.get(5).raw.toString('latin1'), content)
})

// ---------------------------------------------------------------------------
// ToUnicode repair
// ---------------------------------------------------------------------------

test('inspectPdf surfaces malformed ToUnicode blocks', () => {
  const content = 'BT\n/F1 12 Tf\n<0001> Tj\nET'
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' },
    { num: 3, body: '<< /Type /Page /Parent 2 0 R /Contents [ 4 0 R ] /Resources << /Font << /F1 8 0 R >> >> >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${deflate(content).length} >>`, stream: deflate(content) } },
    ...fontObjects({ cmapPairs: CYRILLIC_PAIRS, declaredCount: 944 })
  ]
  const report = inspectPdf(writeTmpPdf(buildPdf(objects)))
  assert.deepEqual(report.malformedToUnicode, [{ object: 10, keyword: 'beginbfchar', declared: 944, actual: 5 }])
})

test('repairToUnicodeCMaps rewrites the count and keeps the file readable', () => {
  const content = 'BT\n/F1 12 Tf\n[<000100040007>] TJ\nET'
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 1 /Kids [ 3 0 R ] >>' },
    { num: 3, body: '<< /Type /Page /Parent 2 0 R /Contents [ 4 0 R ] /Resources << /Font << /F1 8 0 R >> >> >>' },
    { num: 4, body: { dict: `<< /Filter [ /FlateDecode ] /Length ${deflate(content).length} >>`, stream: deflate(content) } },
    ...fontObjects({ cmapPairs: CYRILLIC_PAIRS, declaredCount: 944 })
  ]
  const file = writeTmpPdf(buildPdf(objects))

  const before = inspectPdf(file)
  assert.equal(before.malformedToUnicode.length, 1)

  const result = repairToUnicodeCMaps(file)
  assert.equal(result.repaired, true)
  assert.deepEqual(result.repairs, [{ object: 10, keyword: 'beginbfchar', declared: 944, actual: 5 }])

  const after = inspectPdf(file)
  assert.deepEqual(after.malformedToUnicode, [])
  assert.equal(after.textGlyphs, 3)
  assert.equal(after.cyrillicMappings, 5)
  assert.equal(after.verdict, 'text')

  // The rebuilt cross-reference table must point at real object headers.
  const bytes = fs.readFileSync(file)
  const latin = bytes.toString('latin1')
  const tableMatch = /\nxref\n/.exec(latin)
  assert.ok(tableMatch, 'rebuilt file must contain a cross-reference table')
  const xrefAt = tableMatch.index + 1
  const startxrefAt = latin.lastIndexOf('startxref')
  const startxref = parseInt(latin.slice(startxrefAt + 'startxref'.length).trim(), 10)
  assert.equal(startxref, xrefAt)
  const lines = latin.slice(xrefAt).split('\n')
  assert.equal(lines[0], 'xref')
  let seenEntries = 0
  for (const line of lines.slice(2)) {
    if (!/^\d{10} \d{5} n $/.test(line)) continue
    seenEntries++
    const offset = parseInt(line.slice(0, 10), 10)
    assert.match(latin.slice(offset, offset + 40), /^\d+ 0 obj/, 'xref offset must land on an object header')
  }
  assert.ok(seenEntries >= objects.length, 'every object should have a cross-reference entry')
})

test('repairToUnicodeCMaps is a no-op on a well-formed PDF', () => {
  const objects = [
    { num: 1, body: '<< /Type /Catalog /Pages 2 0 R >>' },
    { num: 2, body: '<< /Type /Pages /Count 0 /Kids [] >>' },
    ...fontObjects({ cmapPairs: CYRILLIC_PAIRS })
  ]
  const file = writeTmpPdf(buildPdf(objects))
  const before = fs.readFileSync(file)
  const result = repairToUnicodeCMaps(file)
  assert.equal(result.repaired, false)
  assert.deepEqual(result.repairs, [])
  assert.deepEqual(fs.readFileSync(file), before, 'an untouched file must not be rewritten')
})
