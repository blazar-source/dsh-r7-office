import fs from 'node:fs'
import path from 'node:path'
import os from 'node:os'
import zlib from 'node:zlib'

/**
 * Test support shared by the PPTX suites.
 *
 * Every fixture a PPTX test needs is authored at run time: the project ships no
 * binary documents, so the images and the chart/SmartArt deck are synthesized
 * here, byte by byte, with the standard library only.
 */

/** A temporary directory that the caller cleans up. */
export function tempDir(label) {
  const dir = path.join(os.tmpdir(), `dsh_r7_${label}_${Date.now()}_${Math.floor(Math.random() * 1e6)}`)
  fs.mkdirSync(dir, { recursive: true })
  return dir
}

const CRC_TABLE = (() => {
  const table = new Int32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c
  }
  return table
})()

function crc32(buffer) {
  let crc = -1
  for (let i = 0; i < buffer.length; i++) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ buffer[i]) & 0xff]
  return (crc ^ -1) >>> 0
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const typeBuffer = Buffer.from(type, 'latin1')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(Buffer.concat([typeBuffer, data])))
  return Buffer.concat([length, typeBuffer, data, crc])
}

/**
 * Encode a real RGB PNG.
 *
 * No image library is available (or allowed: the runtime has zero
 * dependencies), and an invalid PNG would make the "R7 opens the deck" check
 * meaningless, so the encoder is written out here.
 *
 * @param {number} width
 * @param {number} height
 * @param {(x: number, y: number) => [number, number, number]} colourAt
 * @returns {Buffer}
 */
export function makePng(width, height, colourAt) {
  const raw = Buffer.alloc(height * (1 + width * 3))
  let offset = 0
  for (let y = 0; y < height; y++) {
    raw[offset++] = 0 // filter type: none
    for (let x = 0; x < width; x++) {
      const [r, g, b] = colourAt(x, y)
      raw[offset++] = r & 0xff
      raw[offset++] = g & 0xff
      raw[offset++] = b & 0xff
    }
  }

  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = 8 // bit depth
  ihdr[9] = 2 // colour type: truecolour
  ihdr[10] = 0 // deflate
  ihdr[11] = 0 // adaptive filtering
  ihdr[12] = 0 // no interlace

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0))
  ])
}

/** A PNG with a colour gradient, used as the deck's picture. */
export function writeGradientPng(filePath, width = 240, height = 160) {
  const png = makePng(width, height, (x, y) => [
    Math.round((x / Math.max(width - 1, 1)) * 255),
    Math.round((y / Math.max(height - 1, 1)) * 160) + 40,
    200 - Math.round((x / Math.max(width - 1, 1)) * 120)
  ])
  fs.writeFileSync(filePath, png)
  return { filePath, width, height, bytes: png.length }
}

/** A small solid PNG, for replacement tests. */
export function writeSolidPng(filePath, [r, g, b], width = 100, height = 80) {
  const png = makePng(width, height, () => [r, g, b])
  fs.writeFileSync(filePath, png)
  return { filePath, width, height, bytes: png.length }
}

/**
 * Build a file that exercises the parts the engine must never damage.
 *
 * A complete DrawingML chart is added to a real deck: the chart part, the
 * embedded workbook it reads from, the relationship between them, both
 * content-type registrations, the slide relationship and the
 * `<p:graphicFrame>` that shows it. A SmartArt graphic frame follows. That is
 * the shape of a deck whose author used SmartArt, a chart and an embedded
 * workbook — the exact situation the "preserve, never edit" rule is about.
 *
 * The chart is deliberately complete (external data and all) rather than a
 * stub: a half-formed chart makes R7's own converter throw, which would turn a
 * preservation test into a test of my fixture.
 *
 * @param {import('../../src/shared/zip.js').ZipArchive} zip
 * @param {string} slidePartPath
 * @returns {{chartPart: string, chartXml: string, relsPath: string, relId: string}}
 */
export function injectChartAndSmartArt(zip, slidePartPath) {
  const chartPart = 'ppt/charts/chart1.xml'
  const workbookPart = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
  const workbookBuffer = Buffer.from('PK\u0005\u0006' + '\u0000'.repeat(18), 'latin1')

  const chartXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<c:chart><c:plotArea><c:layout/>'
    + '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>'
    + '<c:ser><c:idx val="0"/><c:order val="0"/>'
    + '<c:cat><c:strLit><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat>'
    + '<c:val><c:numLit><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="1"><c:v>19</c:v></c:pt></c:numLit></c:val>'
    + '</c:ser></c:barChart></c:plotArea></c:chart>'
    + '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'
    + '</c:chartSpace>'

  zip.setText(chartPart, chartXml)
  zip.setBuffer(workbookPart, workbookBuffer)
  zip.setText('ppt/charts/_rels/chart1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1"'
    + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"'
    + ' Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/>'
    + '</Relationships>')

  const contentTypes = zip.getText('[Content_Types].xml')
  zip.setText('[Content_Types].xml', contentTypes.replace(
    '</Types>',
    '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
    + '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>'
  ))

  // Relationship from the slide to the chart, and the graphic frames that
  // reference both the chart and an embedded object.
  const relsPath = `ppt/slides/_rels/${path.basename(slidePartPath)}.rels`
  let relsXml = zip.getText(relsPath)
  const relId = `rId${Math.max(0, ...[...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))) + 1}`
  relsXml = relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`
  )
  zip.setText(relsPath, relsXml)

  const chartFrame = '<p:graphicFrame>'
    + '<p:nvGraphicFramePr><p:cNvPr id="700" name="Диаграмма 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    + '<p:xfrm><a:off x="6096000" y="3048000"/><a:ext cx="4876800" cy="2743200"/></p:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    + `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${relId}"/>`
    + '</a:graphicData></a:graphic></p:graphicFrame>'

  const smartArtFrame = '<p:graphicFrame>'
    + '<p:nvGraphicFramePr><p:cNvPr id="701" name="SmartArt 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    + '<p:xfrm><a:off x="609600" y="4572000"/><a:ext cx="3048000" cy="1524000"/></p:xfrm>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
    + '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    + ' r:dm="rId900" r:lo="rId901" r:qs="rId902" r:cs="rId903"/>'
    + '</a:graphicData></a:graphic></p:graphicFrame>'

  const slideXml = zip.getText(slidePartPath)
  zip.setText(slidePartPath, slideXml.replace('</p:spTree>', `${chartFrame}${smartArtFrame}</p:spTree>`))

  return { chartPart, chartXml, relsPath, relId, workbookPart }
}

/** True when every byte of two buffers matches. */
export function sameBytes(a, b) {
  if (!Buffer.isBuffer(a) || !Buffer.isBuffer(b)) return false
  return Buffer.compare(a, b) === 0
}

/**
 * The engine's read model without the private fields the tool layer strips.
 * Tests assert on the public shape, so they use the same projection.
 */
export function publicSlide(slide) {
  return {
    ...slide,
    objects: slide.objects.map(({ _element, _shapeTreeIndex, ...rest }) => rest)
  }
}
